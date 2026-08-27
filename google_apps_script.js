/**
 * 代課費計算與管理系統 - Google Apps Script (GAS) 後台與 Web App 伺服器
 * 【按月分表儲存 (總表 + 明細表) + 內建多頁面 Web App 伺服端】
 *
 * 支援功能：
 * 1. 內建 Web App 網頁伺服：
 *    - 預設首頁（入口門戶）：index.html
 *    - 教師填報端：?page=teacher
 *    - 行政核銷端：?page=admin
 * 2. 支援 google.script.run 直連（免除 CORS 與網址設定困擾）。
 * 3. 相容既有 HTTP API (GET / POST) 與 Python 核銷腳本。
 */

/**
/**
 * 預設安全金鑰（系統設定分頁未設定時之預設值）
 * ⚠️ 現在可在「系統全域設定」介面中直接修改與儲存此金鑰，不再需要手動修改程式碼。
 */
var DEFAULT_SECRET_KEY = '087525402';

// 課表分頁名稱（每學年更新課表時，直接覆蓋這個分頁的內容即可，不必改程式碼）
var TIMETABLE_SHEET = 'Class_Timetables';

// 系統設定分頁名稱（集中管理全校一致性設定項目）
var SETTINGS_SHEET = 'System_Settings';
var SETTINGS_HEADERS = ["項目名稱(Name)", "設定代碼(Key)", "設定值(Value)", "說明(Description)"];

var DEFAULT_SETTINGS = [
  { name: "學校名稱", key: "school_name", value: "馬鳴國小", desc: "系統全銜與印領清冊抬頭" },
  { name: "系統維護人員職稱", key: "maintainer_title", value: "資訊執秘", desc: "負責系統運作維護/網管人員職稱" },
  { name: "系統維護人員姓名", key: "maintainer_name", value: "蔡志益", desc: "負責系統運作維護/網管人員姓名" },
  { name: "教育階段", key: "school_level", value: "elementary", desc: "elementary(國小) / junior(國中) / senior(高中)" },
  { name: "單節鐘點費率", key: "rate_per_period", value: "405", desc: "國小 405 元/節 (依教育部函文)" },
  { name: "學年學期", key: "academic_year_term", value: "114-1", desc: "當前運行的學年與學期" },
  { name: "代理導師預設日薪", key: "mentor_daily_rate", value: "1528", desc: "整天代導師之日薪預設標準" },
  { name: "非計費作息項目", key: "duty_items", value: "早修,打掃,午餐,午休,放學", desc: "交接單上之作息指導項目(逗號分隔)" },
  { name: "安全管理金鑰", key: "secret_key", value: "087525402", desc: "行政端結算與課表匯入之安全管理金鑰(密碼)" }
];

var SUMMARY_HEADERS = [
  "申請編號(ID)", "填報時間(Timestamp)", "代課日期(Date)", "學校名稱(School)",
  "教育階段(Level)", "請假教師(AbsentTeacher)", "代課教師(SubTeacher)", "班級(ClassName)",
  "假別(LeaveType)", "計費節數(PeriodsCount)", "單節費率(Rate)", "課堂鐘點費(ClassFee)",
  "是否代理導師(IsActingMentor)", "代理導師天數(ActingDays)", "導師費加給(MentorFee)",
  "整單總金額(TotalFee)", "計費模式(PayMode)", "備註(PayNote)"
];

var DETAIL_HEADERS = [
  "申請編號(ID)", "代課日期(Date)", "班級(ClassName)", "節次(Period)",
  "科目(Subject)", "代課教師(SubTeacher)", "是否計費(Paid)", "單節費率(Rate)",
  "預估金額(Fee)", "核銷狀態(Status)", "交代事項(Handover)", "確認日期(ConfirmDate)"
];

// 明細表欄位位置（1-indexed，供狀態更新使用）
var DETAIL_COL_STATUS = 10;
var DETAIL_COL_CONFIRM_DATE = 12;

/**
 * 安全取得目前試算表
 */
function getActiveSs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("無法連接 Google 試算表。請確認此 Apps Script 是在 Google 試算表內點選「擴充功能」➔「Apps Script」建立的。");
  }
  return ss;
}

/**
 * 容錯樣板載入器（自動嘗試 filename, filename.html）
 */
function createTemplateHelper(filename) {
  var cleanName = filename.replace(/\.html$/, '');
  var candidates = [cleanName, cleanName + '.html', filename];
  for (var i = 0; i < candidates.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(candidates[i]);
    } catch (e) {
      // 繼續嘗試下一個候選檔名
    }
  }
  throw new Error("找不到名為「" + filename + "」的 HTML 檔案，請確認 Apps Script 專案中已建立該檔案。");
}

/**
 * HTML 樣式/腳本引入輔助函式（供 HTML 樣板內 <?!= include('styles'); ?> 使用）
 */
function include(filename) {
  var cleanName = filename.replace(/\.html$/, '');
  var candidates = [cleanName, cleanName + '.html', filename];
  for (var i = 0; i < candidates.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(candidates[i]).getContent();
    } catch (e) {
      // 繼續嘗試下一個候選檔名
    }
  }
  return '<style>/* 未找到 ' + filename + ' 樣式檔案 */</style>';
}

/**
 * 取得目前 Web App 部署網址（自動移除教育網域前綴，防止多帳號登入跳轉衝突）
 */
function getScriptUrl() {
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) {
      // 自動轉換 /a/網域名稱/ 為標準公開網址
      url = url.replace(/\/a\/[^\/]+\/macros\//, '/macros/');
    }
    return url || '';
  } catch (e) {
    return '';
  }
}

/**
 * 輔助：依頁面名稱與學校名稱動態取得分頁標題
 */
function getTitleForPage(page, schoolName) {
  var school = (schoolName && String(schoolName).trim()) ? String(schoolName).trim() : '代課系統';
  if (page === 'admin') return school + ' - 代課費計算與管理後台 (行政端)';
  if (page === 'teacher') return school + ' - 代課與交接單填報 (教師端)';
  return school + ' - 代課費計算與管理系統';
}

/**
 * 處理 Web 頁面渲染與 GET API 請求
 */
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};

    // -------------------------------------------------------------
    // 1. API 模式：供外部工具 (如 Python、第三方請求) 取得資料
    // -------------------------------------------------------------
    if (params.action === "timetable") {
      var ss = getActiveSs();
      return getCorsResponse(JSON.stringify(buildTimetable(ss)), 'json');
    }

    if (params.action === "settings") {
      var ss = getActiveSs();
      return getCorsResponse(JSON.stringify(getSystemSettings(ss)), 'json');
    }

    if (params.month || params.action === "confirm") {
      var ss = getActiveSs();
      // 驗證金鑰
      if (!verifyAuthKey(params.key, ss)) {
        return getCorsResponse(JSON.stringify({
          status: "error",
          message: "未授權：缺少或錯誤的金鑰。請在系統全域設定或右上角鑰匙圖示中輸入正確金鑰。"
        }), 'json');
      }

      // 狀態修改 (action = confirm)
      if (params.action === "confirm" && params.month && params.appId && params.period) {
        var confirmResult = confirmPeriodRecord(params.month, params.appId, params.period, params.key);
        return getCorsResponse(JSON.stringify(confirmResult), 'json');
      }

      // 讀取全部月份 (month = all)
      if (params.month === "all") {
        return getCorsResponse(JSON.stringify(getAllRecordsData(params.key)), 'json');
      }

      // 讀取單一月份
      var targetMonth = params.month;
      if (!targetMonth) {
        var today = new Date();
        targetMonth = today.getFullYear() + "-" + ("0" + (today.getMonth() + 1)).slice(-2);
      }
      return getCorsResponse(JSON.stringify(buildRecordsForMonth(ss, targetMonth)), 'json');
    }

    // -------------------------------------------------------------
    // 2. 網頁渲染模式 (Google Apps Script Web App 頁面切換)
    // -------------------------------------------------------------
    var page = params.page || 'index';
    if (page !== 'admin' && page !== 'teacher' && page !== 'index') {
      page = 'index';
    }

    var template = createTemplateHelper(page);
    template.page = page;
    template.scriptUrl = getScriptUrl();

    var activeSs = null;
    try {
      activeSs = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {}

    var currentSettings = activeSs ? getSystemSettings(activeSs) : {};
    template.systemSettings = currentSettings;
    var currentSchool = (currentSettings && currentSettings.school_name) ? currentSettings.school_name : '代課系統';
    var mTitle = (currentSettings && (currentSettings.maintainer_title || currentSettings.admin_title)) ? (currentSettings.maintainer_title || currentSettings.admin_title) : '系統維護';
    var mName = (currentSettings && (currentSettings.maintainer_name || currentSettings.admin_name)) ? (currentSettings.maintainer_name || currentSettings.admin_name) : '';
    template.currentSchool = currentSchool;
    template.currentMaintainer = (mTitle + ' ' + mName).trim();

    if (page === 'teacher') {
      try {
        template.serverTimetable = activeSs ? buildTimetable(activeSs) : {};
      } catch (e) {
        template.serverTimetable = {};
      }
    }

    return template.evaluate()
      .setTitle(getTitleForPage(page, currentSchool))
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    var errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>系統載入錯誤</title><style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:2rem;background:#f8fafc;color:#1e293b;} .error-card{background:white;padding:2rem;border-radius:12px;border:1px solid #fee2e2;box-shadow:0 4px 6px -1px rgba(0,0,0,0.1);max-width:600px;margin:2rem auto;} h2{color:#dc2626;display:flex;align-items:center;gap:0.5rem;} pre{background:#f1f5f9;padding:1rem;border-radius:6px;overflow-x:auto;color:#b91c1c;font-size:0.9rem;} ol{margin-left:1.5rem;line-height:1.8;}</style></head><body><div class="error-card"><h2>⚠️ 系統載入發生問題</h2><p>錯誤訊息：</p><pre>' + error.toString() + '</pre><hr style="margin:1.5rem 0;border:none;border-top:1px solid #e2e8f0;"><p><strong>常見排查方式：</strong></p><ol><li>請確認 Google Apps Script 專案中已建立 <code>Code.gs</code>、<code>index.html</code>、<code>teacher.html</code>、<code>admin.html</code> 等檔案。</li><li>請確認 Google 試算表內已建立 <code>Class_Timetables</code> 分頁並填入課表。</li><li>重新部署時請點選「部署」➔「管理部署」➔「編輯」➔版本選「新版本」後儲存。</li></ol></div></body></html>';
    return HtmlService.createHtmlOutput(errorHtml);
  }
}

/**
 * 接收外部 POST 請求 (相容傳統 fetch POST)
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var result = saveHandoverRecord(data);
    return getCorsResponse(JSON.stringify(result), 'json');
  } catch (error) {
    return getCorsResponse(JSON.stringify({ status: "error", message: error.toString() }), 'json');
  }
}

// =================================================================
// 伺服端 RPC 函式 (供網頁端 google.script.run 直連呼叫)
// =================================================================

/**
 * 儲存教師交接單與代課紀錄（伺服端直連）
 */
function saveHandoverRecord(data) {
  try {
    var dateStr = data.date; // 格式為 YYYY-MM-DD
    if (!dateStr) {
      return { status: "error", message: "缺少代課日期 (date)！" };
    }

    var month = dateStr.substring(0, 7); // YYYY-MM
    var ss = getActiveSs();
    var summarySheet = getOrCreateSheet(ss, month + "_總表", SUMMARY_HEADERS);
    var detailSheet = getOrCreateSheet(ss, month + "_明細", DETAIL_HEADERS);

    var rate = Number(data.rate) || 0;
    var payMode = data.payMode || 'perPeriod';
    var isActingMentor = data.isActingMentor === true || payMode === 'mentorDaily';

    var appId = "SUB-" + month.replace("-", "") + "-" + Math.floor(100 + Math.random() * 900);
    summarySheet.appendRow([
      appId,
      new Date().toISOString(),
      data.date,
      data.school || "",
      data.level || "",
      data.absentTeacher || "",
      data.subTeacher || "",
      data.className || "",
      data.leaveType || "",
      Number(data.periodsCount) || 0,
      rate,
      Number(data.classFee) || 0,
      isActingMentor ? "是" : "否",
      Number(data.actingDays) || 0,
      Number(data.mentorFee) || 0,
      Number(data.totalFee) || 0,
      payMode,
      data.payNote || ""
    ]);

    // 明細：一次性批次寫入 (Batch I/O，效能提升 3~5 倍)
    if (data.periodsDetail && Array.isArray(data.periodsDetail) && data.periodsDetail.length > 0) {
      var detailRows = [];
      for (var i = 0; i < data.periodsDetail.length; i++) {
        var p = data.periodsDetail[i];
        var isPaid = p.paid === true;
        var periodRate = isPaid ? rate : 0;

        detailRows.push([
          appId,
          data.date,
          data.className || "",
          p.period,
          p.subject || "",
          p.subTeacher || data.subTeacher || "",
          isPaid ? "是" : "否",
          periodRate,
          periodRate,
          "待核對",
          p.handover || "",
          ""
        ]);
      }
      if (detailRows.length > 0) {
        var startRow = detailSheet.getLastRow() + 1;
        detailSheet.getRange(startRow, 1, detailRows.length, DETAIL_HEADERS.length).setValues(detailRows);
      }
    }

    return { status: "success", appId: appId, message: "資料已成功按月寫入 Google 試算表！" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

/**
 * 批次儲存多日連續交接單紀錄（伺服端直連，支援一次性寫入多天資料）
 */
function saveBatchHandoverRecords(recordsList) {
  if (!Array.isArray(recordsList) || recordsList.length === 0) {
    return { status: "error", message: "批次清單為空！" };
  }
  try {
    var ss = getActiveSs();
    var successCount = 0;
    for (var r = 0; r < recordsList.length; r++) {
      var item = recordsList[r];
      var res = saveHandoverRecord(item);
      if (res && res.status === "success") {
        successCount++;
      }
    }
    return {
      status: "success",
      count: successCount,
      message: "成功批次寫入 " + successCount + " 天代課交接紀錄至 Google 試算表！"
    };
  } catch (err) {
    return { status: "error", message: "批次儲存失敗：" + err.toString() };
  }
}

/**
 * 取得學校課表資料庫（伺服端直連）
 */
function getTimetableData() {
  var ss = getActiveSs();
  return buildTimetable(ss);
}

/**
 * 取得全域系統設定（伺服端直連）
 */
function getSystemSettingsData() {
  var ss = getActiveSs();
  return getSystemSettings(ss);
}

/**
 * 儲存全域系統設定（需驗證金鑰，伺服端直連）
 */
function saveSystemSettingsData(settingsObj, key) {
  return saveSystemSettings(settingsObj, key);
}

/**
 * 取得當前有效的安全管理金鑰（優先由 System_Settings 讀取，若無則使用預設值）
 */
function getValidSecretKey(ss) {
  if (!ss) ss = getActiveSs();
  if (!ss) return DEFAULT_SECRET_KEY;
  try {
    var sheet = ss.getSheetByName(SETTINGS_SHEET);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var displayValues = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
        for (var i = 0; i < displayValues.length; i++) {
          if (String(displayValues[i][1] || '').trim() === 'secret_key') {
            var val = displayValues[i][2];
            if (val !== null && val !== undefined && String(val).trim() !== '') {
              var kStr = String(val).trim();
              if (kStr === '87525402') kStr = '087525402';
              return kStr;
            }
          }
        }
      }
    }
  } catch (e) {}
  return DEFAULT_SECRET_KEY;
}

/**
 * 驗證傳入的金鑰是否正確
 */
function verifyAuthKey(providedKey, ss) {
  var validKey = getValidSecretKey(ss);
  return (providedKey !== null && providedKey !== undefined && String(providedKey).trim() === validKey);
}

/**
 * 批次匯入課表資料（需驗證金鑰，伺服端直連）
 * 支援三種模式：
 * 1. 'smart_upsert' (預設推薦): 智慧更新涉及班級，其餘班級維持不變
 * 2. 'append': 增量追加新紀錄
 * 3. 'overwrite': 完全覆蓋全校課表 (執行前自動建立備份分頁)
 */
function importTimetableData(rows, mode, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    return { status: "error", message: "未授權：安全金鑰錯誤或未填寫。請在系統全域設定中輸入正確金鑰。" };
  }
  if (!rows || !Array.isArray(rows) || rows.length === 0) {
    return { status: "error", message: "匯入資料為空，請確認格式！" };
  }
  try {
    var sheet = ss.getSheetByName(TIMETABLE_SHEET);
    var TIMETABLE_HEADERS = ["班級(ClassName)", "星期(DayOfWeek)", "節次(Period)", "科目(Subject)", "授課教師(Teacher)"];
    
    if (!sheet) {
      sheet = ss.insertSheet(TIMETABLE_SHEET);
      sheet.appendRow(TIMETABLE_HEADERS);
      sheet.setFrozenRows(1);
    }

    var dataToWrite = [];
    var incomingClassesMap = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var cls = String(r[0] !== null && r[0] !== undefined ? r[0] : '').trim();
      var day = parseInt(r[1], 10);
      var period = parseInt(r[2], 10);
      var subject = String(r[3] || '').trim();
      var teacher = String(r[4] || '').trim();

      if (!cls || isNaN(day) || isNaN(period)) continue;
      dataToWrite.push([cls, day, period, subject, teacher]);
      incomingClassesMap[cls] = true;
    }

    if (dataToWrite.length === 0) {
      return { status: "error", message: "沒有合法的課表資料（請檢查是否有包含班級、星期1~5、節次1~7）！" };
    }

    var incomingClassesList = Object.keys(incomingClassesMap);
    var responseMessage = "";

    if (mode === 'overwrite') {
      // 模式 3: 完全覆蓋全校課表 (高風險操作，自動建立備份分頁)
      var oldLastRow = sheet.getLastRow();
      var backupSheetName = "";
      if (oldLastRow > 1) {
        try {
          var dateStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmmss");
          backupSheetName = "課表備份_" + dateStr;
          var backupSheet = ss.insertSheet(backupSheetName);
          var oldValues = sheet.getRange(1, 1, oldLastRow, 5).getValues();
          backupSheet.getRange(1, 1, oldValues.length, 5).setValues(oldValues);
          backupSheet.setFrozenRows(1);
        } catch (bErr) {
          console.log("自動備份分頁略過: " + bErr);
        }
        sheet.getRange(2, 1, oldLastRow - 1, 5).clearContent();
      }

      sheet.getRange(2, 1, dataToWrite.length, 5).setValues(dataToWrite);
      responseMessage = "【全校覆蓋成功】已完全覆蓋全校課表，共寫入 " + dataToWrite.length + " 筆！" +
        (backupSheetName ? "（原課表已自動安全備份至「" + backupSheetName + "」分頁）" : "");

    } else if (mode === 'append') {
      // 模式 2: 純增量追加
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, dataToWrite.length, 5).setValues(dataToWrite);
      responseMessage = "【增量追加成功】已成功追加 " + dataToWrite.length + " 筆課表紀錄！";

    } else {
      // 模式 1 (預設): 智慧班級更新 (Smart Upsert)
      // 讀取既有資料，只清除本次有出現的班級，其餘班級保留
      var currentLastRow = sheet.getLastRow();
      var retainedRows = [];
      if (currentLastRow > 1) {
        var currentValues = sheet.getRange(2, 1, currentLastRow - 1, 5).getValues();
        for (var k = 0; k < currentValues.length; k++) {
          var existCls = String(currentValues[k][0] || '').trim();
          if (existCls && !incomingClassesMap[existCls]) {
            retainedRows.push(currentValues[k]);
          }
        }
      }

      var finalData = retainedRows.concat(dataToWrite);

      // 清空舊內容並一次寫入整併後的完整課表
      if (currentLastRow > 1) {
        sheet.getRange(2, 1, currentLastRow - 1, 5).clearContent();
      }
      sheet.getRange(2, 1, finalData.length, 5).setValues(finalData);

      responseMessage = "【智慧更新成功】已更新「" + incomingClassesList.join("、") + "」等 " +
        incomingClassesList.length + " 個班級課表（共 " + dataToWrite.length + " 節），其餘班級完整保留！目前全校共 " +
        finalData.length + " 節課。";
    }

    // 清除課表快取以使新匯入資料即時生效
    try {
      CacheService.getScriptCache().remove("TIMETABLE_DB");
    } catch (cErr) {}

    return {
      status: "success",
      count: dataToWrite.length,
      message: responseMessage,
      timetable: buildTimetable(ss)
    };
  } catch (err) {
    return { status: "error", message: "匯入發生錯誤：" + err.toString() };
  }
}

/**
 * 讀取系統設定底層函式（支援 CacheService 1小時高效快取）
 */
function getSystemSettings(ss) {
  if (!ss) return {};
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("SYSTEM_SETTINGS");
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (cErr) {}

  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.appendRow(SETTINGS_HEADERS);
    for (var i = 0; i < DEFAULT_SETTINGS.length; i++) {
      var item = DEFAULT_SETTINGS[i];
      sheet.appendRow([item.name, item.key, item.value, item.desc]);
    }
    sheet.setFrozenRows(1);
    sheet.getRange(2, 3, DEFAULT_SETTINGS.length, 1).setNumberFormat('@');
  }

  var lastRow = sheet.getLastRow();
  var settings = {};
  for (var k = 0; k < DEFAULT_SETTINGS.length; k++) {
    settings[DEFAULT_SETTINGS[k].key] = DEFAULT_SETTINGS[k].value;
  }

  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).setNumberFormat('@');
    var displayValues = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
    for (var j = 0; j < displayValues.length; j++) {
      var kName = String(displayValues[j][1] || '').trim();
      var val = displayValues[j][2];
      if (kName) {
        var strVal = (val !== null && val !== undefined) ? String(val).trim() : '';
        if (kName === 'secret_key' && strVal === '87525402') {
          strVal = '087525402';
        }
        settings[kName] = strVal;
      }
    }
  }

  // 雙向補全維護人員與舊版欄位相容
  if (settings['maintainer_title'] && !settings['admin_title']) {
    settings['admin_title'] = settings['maintainer_title'];
  }
  if (settings['admin_title'] && !settings['maintainer_title']) {
    settings['maintainer_title'] = settings['admin_title'];
  }
  if (settings['maintainer_name'] && !settings['admin_name']) {
    settings['admin_name'] = settings['maintainer_name'];
  }
  if (settings['admin_name'] && !settings['maintainer_name']) {
    settings['maintainer_name'] = settings['admin_name'];
  }

  try {
    CacheService.getScriptCache().put("SYSTEM_SETTINGS", JSON.stringify(settings), 300);
  } catch (pErr) {}

  return settings;
}

/**
 * 儲存系統設定底層函式
 */
function saveSystemSettings(settingsObj, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    return { status: "error", message: "未授權：安全金鑰錯誤或未填寫。請輸入正確金鑰。" };
  }
  try {
    var sheet = ss.getSheetByName(SETTINGS_SHEET);
    if (!sheet) {
      getSystemSettings(ss);
      sheet = ss.getSheetByName(SETTINGS_SHEET);
    }
    
    var lastRow = sheet.getLastRow();
    var existingKeys = {};
    if (lastRow > 1) {
      var displayValues = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
      for (var i = 0; i < displayValues.length; i++) {
        var kName = String(displayValues[i][1] || '').trim();
        if (kName) existingKeys[kName] = i + 2;
      }
    }

    sheet.getRange(2, 3, Math.max(1, lastRow), 1).setNumberFormat('@');

    for (var skey in settingsObj) {
      var sval = String(settingsObj[skey] || '').trim();
      if (existingKeys[skey]) {
        sheet.getRange(existingKeys[skey], 3).setNumberFormat('@').setValue(sval);
      } else {
        var sname = skey;
        var sdesc = "";
        for (var d = 0; d < DEFAULT_SETTINGS.length; d++) {
          if (DEFAULT_SETTINGS[d].key === skey) {
            sname = DEFAULT_SETTINGS[d].name;
            sdesc = DEFAULT_SETTINGS[d].desc;
            break;
          }
        }
        sheet.appendRow([sname, skey, sval, sdesc]);
        var newRow = sheet.getLastRow();
        sheet.getRange(newRow, 3).setNumberFormat('@').setValue(sval);
      }
    }

    // 清除快取以使新設定即刻生效
    try {
      CacheService.getScriptCache().remove("SYSTEM_SETTINGS");
    } catch (rErr) {}

    return { status: "success", message: "全域系統設定已成功儲存至 Google 試算表！", settings: getSystemSettings(ss) };
  } catch (err) {
    return { status: "error", message: "儲存失敗：" + err.toString() };
  }
}

/**
 * 取得指定月份的代課紀錄（需驗證金鑰，伺服端直連）
 */
function getMonthlyRecords(targetMonth, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    throw new Error("未授權：金鑰錯誤或未填寫。請至系統全域設定或右上角鑰匙圖示中填入正確金鑰。");
  }
  return buildRecordsForMonth(ss, targetMonth);
}

/**
 * 取得所有歷史月份的代課紀錄（需驗證金鑰，伺服端直連）
 */
function getAllRecordsData(key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    throw new Error("未授權：金鑰錯誤或未填寫。請至系統全域設定或右上角鑰匙圖示中填入正確金鑰。");
  }
  var months = listDataMonths(ss);
  var allRecords = [];
  for (var n = 0; n < months.length; n++) {
    allRecords = allRecords.concat(buildRecordsForMonth(ss, months[n]));
  }
  allRecords.sort(function (a, b) {
    return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0);
  });
  return allRecords;
}

/**
 * 更新明細核銷確認狀態（需驗證金鑰，伺服端直連）
 */
function confirmPeriodRecord(month, appId, period, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    return { status: "error", message: "未授權：金鑰錯誤。" };
  }
  var ss = getActiveSs();
  var detailSheet = ss.getSheetByName(month + "_明細");
  if (!detailSheet) {
    return { status: "error", message: "找不到該月份的明細分頁！" };
  }

  var lastRow = detailSheet.getLastRow();
  if (lastRow <= 1) {
    return { status: "error", message: "該月份明細分頁尚無資料！" };
  }

  var values = detailSheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues();
  var success = false;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === appId && values[i][3].toString() === period.toString()) {
      var rowNum = i + 2;
      detailSheet.getRange(rowNum, DETAIL_COL_STATUS).setValue("已確認");
      detailSheet.getRange(rowNum, DETAIL_COL_CONFIRM_DATE).setValue(new Date().toISOString());
      success = true;
      break;
    }
  }

  return success
    ? { status: "success", message: "狀態已更新為已確認！" }
    : { status: "error", message: "找不到對應的代課節次明細！" };
}

// =================================================================
// 試算表底層資料處理函式
// =================================================================

function getCorsResponse(content, type) {
  var output = ContentService.createTextOutput(content);
  if (type === 'json') {
    output.setMimeType(ContentService.MimeType.JSON);
  } else {
    output.setMimeType(ContentService.MimeType.TEXT);
  }
  return output;
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function buildTimetable(ss) {
  if (!ss) return { db: {}, teachers: [], classes: [] };
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("TIMETABLE_DB");
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (cErr) {}

  var sheet = ss.getSheetByName(TIMETABLE_SHEET);
  if (!sheet) return { db: {}, teachers: [], classes: [] };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { db: {}, teachers: [], classes: [] };

  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var db = {};
  var teachersMap = {};
  var classesMap = {};

  for (var i = 0; i < values.length; i++) {
    var rawCls = values[i][0];
    if (rawCls === null || rawCls === undefined) continue;
    var cls = String(rawCls).trim();
    var day = parseInt(values[i][1], 10);
    var period = parseInt(values[i][2], 10);
    var subject = String(values[i][3] || '').trim();
    var teacher = String(values[i][4] || '').trim();
    if (!cls || isNaN(day) || isNaN(period)) continue;

    if (!db[cls]) db[cls] = {};
    if (!db[cls][day]) db[cls][day] = {};
    db[cls][day][period] = { subject: subject, teacher: teacher };

    if (cls) classesMap[cls] = true;
    if (teacher) teachersMap[teacher] = true;
  }

  var teachers = Object.keys(teachersMap).sort();
  var classes = Object.keys(classesMap).sort();
  var result = { db: db, teachers: teachers, classes: classes };

  try {
    CacheService.getScriptCache().put("TIMETABLE_DB", JSON.stringify(result), 3600);
  } catch (pErr) {}

  return result;
}

function listDataMonths(ss) {
  var months = [];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(/^(\d{4}-\d{2})_總表$/);
    if (m) months.push(m[1]);
  }
  months.sort();
  return months;
}

function buildRecordsForMonth(ss, targetMonth) {
  var summarySheet = ss.getSheetByName(targetMonth + "_總表");
  var detailSheet = ss.getSheetByName(targetMonth + "_明細");
  if (!summarySheet || !detailSheet) return [];

  var sumLastRow = summarySheet.getLastRow();
  var detLastRow = detailSheet.getLastRow();
  if (sumLastRow <= 1) return [];

  var sumValues = summarySheet.getRange(2, 1, sumLastRow - 1, SUMMARY_HEADERS.length).getValues();
  var detValues = detLastRow > 1
    ? detailSheet.getRange(2, 1, detLastRow - 1, DETAIL_HEADERS.length).getValues()
    : [];

  var detailsMap = {};
  for (var j = 0; j < detValues.length; j++) {
    var d = detValues[j];
    var key = d[0];
    if (!detailsMap[key]) detailsMap[key] = [];
    detailsMap[key].push({
      period: d[3].toString(),
      className: d[2],
      subject: d[4],
      subTeacher: d[5],
      paid: d[6] === "是",
      rate: Number(d[7]) || 0,
      fee: Number(d[8]) || 0,
      status: d[9],
      handover: d[10],
      confirmDate: d[11]
    });
  }

  var records = [];
  for (var k = 0; k < sumValues.length; k++) {
    var s = sumValues[k];
    var currentAppId = s[0];
    var periodsCount = Number(s[9]) || 0;
    var rate = Number(s[10]) || 0;
    var classFee = Number(s[11]) || (periodsCount * rate);
    var isActingMentor = s[12] === "是";
    var actingDays = Number(s[13]) || 0;
    var mentorFee = Number(s[14]) || 0;
    var totalFee = Number(s[15]) || (classFee + mentorFee);

    records.push({
      id: currentAppId,
      timestamp: s[1],
      date: formatDate(s[2]),
      school: s[3],
      level: s[4],
      absentTeacher: s[5],
      subTeacher: s[6],
      className: s[7],
      leaveType: s[8],
      periodsCount: periodsCount,
      rate: rate,
      classFee: classFee,
      isActingMentor: isActingMentor,
      actingDays: actingDays,
      mentorFee: mentorFee,
      totalFee: totalFee,
      payMode: s[16] || 'perPeriod',
      payNote: s[17] || "",
      periodsDetail: detailsMap[currentAppId] || []
    });
  }
  return records;
}

/**
 * 更新指定月份各代課教師之導師加給與總額（需驗證金鑰，伺服端直連）
 * @param {string} targetMonth - 例如 "2026-08"
 * @param {Object} teacherMentorFeesMap - 例如 { "王小明": 1528, "蔡志益": 0 }
 * @param {string} key - 安全管理金鑰
 */
function updateMonthlyMentorFees(targetMonth, teacherMentorFeesMap, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    return { status: "error", message: "未授權：安全金鑰錯誤。" };
  }
  try {
    var summarySheet = ss.getSheetByName(targetMonth + "_總表");
    if (!summarySheet) {
      return { status: "error", message: "找不到該月份的總表分頁！" };
    }
    var lastRow = summarySheet.getLastRow();
    if (lastRow <= 1) {
      return { status: "error", message: "該月份總表尚無紀錄！" };
    }

    var values = summarySheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).getValues();
    var updatedTeachers = {};
    for (var i = 0; i < values.length; i++) {
      var teacher = String(values[i][6] || '').trim();
      var periods = Number(values[i][9]) || 0;
      var rate = Number(values[i][10]) || 0;
      var classFee = Number(values[i][11]) || (periods * rate);

      if (teacherMentorFeesMap && teacherMentorFeesMap.hasOwnProperty(teacher)) {
        var currentMentorFee = 0;
        if (!updatedTeachers[teacher]) {
          currentMentorFee = Number(teacherMentorFeesMap[teacher]) || 0;
          updatedTeachers[teacher] = true;
        } else {
          currentMentorFee = 0;
        }
        var totalFee = classFee + currentMentorFee;
        var rowNum = i + 2;
        summarySheet.getRange(rowNum, 12).setValue(classFee); // 課堂鐘點費
        summarySheet.getRange(rowNum, 15).setValue(currentMentorFee); // 導師費加給
        summarySheet.getRange(rowNum, 16).setValue(totalFee); // 整單總金額
        if (currentMentorFee > 0) {
          summarySheet.getRange(rowNum, 13).setValue("是");
        }
      }
    }

    return {
      status: "success",
      message: "🎉 " + targetMonth + " 導師加給與應領總額已成功儲存至 Google 試算表！",
      records: buildRecordsForMonth(ss, targetMonth)
    };
  } catch (err) {
    return { status: "error", message: "更新導師加給失敗：" + err.toString() };
  }
}

/**
 * 批次更新指定明細核銷狀態（需驗證金鑰，伺服端直連）
 */
function batchUpdateRecordStatus(targetMonth, appIds, newStatus, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    return { status: "error", message: "未授權：安全金鑰錯誤。" };
  }
  try {
    var detailSheet = ss.getSheetByName(targetMonth + "_明細");
    if (!detailSheet) {
      return { status: "error", message: "找不到該月份的明細分頁！" };
    }
    var lastRow = detailSheet.getLastRow();
    if (lastRow <= 1) {
      return { status: "error", message: "該月份明細分頁尚無資料！" };
    }

    var values = detailSheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues();
    var count = 0;
    var nowIso = new Date().toISOString();

    for (var i = 0; i < values.length; i++) {
      var id = String(values[i][0] || '');
      if (!appIds || appIds.length === 0 || appIds.indexOf(id) !== -1) {
        var rowNum = i + 2;
        detailSheet.getRange(rowNum, DETAIL_COL_STATUS).setValue(newStatus);
        if (newStatus === "已核銷" || newStatus === "已確認") {
          detailSheet.getRange(rowNum, DETAIL_COL_CONFIRM_DATE).setValue(nowIso);
        } else {
          detailSheet.getRange(rowNum, DETAIL_COL_CONFIRM_DATE).setValue("");
        }
        count++;
      }
    }

    return {
      status: "success",
      message: "已更新 " + count + " 筆明細狀態為「" + newStatus + "」！",
      records: buildRecordsForMonth(ss, targetMonth)
    };
  } catch (err) {
    return { status: "error", message: "批次更新失敗：" + err.toString() };
  }
}

/**
 * 刪除單筆代課紀錄（同步自總表與明細分頁中清除）
 */
function deleteMonthlyRecord(targetMonth, appId, key) {
  return batchDeleteMonthlyRecords(targetMonth, [appId], key);
}

/**
 * 批次刪除指定代課紀錄（需驗證金鑰，伺服端直連）
 */
function batchDeleteMonthlyRecords(targetMonth, appIds, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    return { status: "error", message: "未授權：安全金鑰錯誤。" };
  }
  if (!appIds || !Array.isArray(appIds) || appIds.length === 0) {
    return { status: "error", message: "請指定要刪除的代課單號！" };
  }
  try {
    var summarySheet = ss.getSheetByName(targetMonth + "_總表");
    var detailSheet = ss.getSheetByName(targetMonth + "_明細");
    var deletedCount = 0;

    // 1. 從總表中由下往上刪除對應列
    if (summarySheet) {
      var sumLastRow = summarySheet.getLastRow();
      if (sumLastRow > 1) {
        var sumValues = summarySheet.getRange(2, 1, sumLastRow - 1, 1).getValues();
        for (var i = sumValues.length - 1; i >= 0; i--) {
          var id = String(sumValues[i][0] || '').trim();
          if (appIds.indexOf(id) !== -1) {
            summarySheet.deleteRow(i + 2);
            deletedCount++;
          }
        }
      }
    }

    // 2. 從明細表中由下往上刪除對應列
    if (detailSheet) {
      var detLastRow = detailSheet.getLastRow();
      if (detLastRow > 1) {
        var detValues = detailSheet.getRange(2, 1, detLastRow - 1, 1).getValues();
        for (var j = detValues.length - 1; j >= 0; j--) {
          var did = String(detValues[j][0] || '').trim();
          if (appIds.indexOf(did) !== -1) {
            detailSheet.deleteRow(j + 2);
          }
        }
      }
    }

    return {
      status: "success",
      message: "🎉 已成功刪除 " + deletedCount + " 筆代課單資料！",
      records: buildRecordsForMonth(ss, targetMonth)
    };
  } catch (err) {
    return { status: "error", message: "刪除失敗：" + err.toString() };
  }
}

/**
 * 線上修改單筆代課紀錄（同步更新總表與明細）
 */
function updateMonthlyRecord(targetMonth, appId, updatedData, key) {
  var ss = getActiveSs();
  if (!verifyAuthKey(key, ss)) {
    return { status: "error", message: "未授權：安全金鑰錯誤。" };
  }
  if (!appId || !updatedData) {
    return { status: "error", message: "缺少更新資料或單號！" };
  }
  try {
    var summarySheet = ss.getSheetByName(targetMonth + "_總表");
    var detailSheet = ss.getSheetByName(targetMonth + "_明細");
    if (!summarySheet) {
      return { status: "error", message: "找不到該月份的總表！" };
    }

    var sumLastRow = summarySheet.getLastRow();
    if (sumLastRow <= 1) {
      return { status: "error", message: "總表尚無資料！" };
    }

    var sumValues = summarySheet.getRange(2, 1, sumLastRow - 1, SUMMARY_HEADERS.length).getValues();
    var targetRowIndex = -1;
    for (var i = 0; i < sumValues.length; i++) {
      if (String(sumValues[i][0] || '').trim() === String(appId).trim()) {
        targetRowIndex = i + 2;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { status: "error", message: "在總表中找不到單號 " + appId };
    }

    var date = updatedData.date;
    var absentTeacher = updatedData.absentTeacher || '';
    var subTeacher = updatedData.subTeacher || '';
    var className = updatedData.className || '';
    var leaveType = updatedData.leaveType || '';
    var periodsCount = Number(updatedData.periodsCount) || 0;
    var rate = Number(updatedData.rate) || 405;
    var classFee = Number(updatedData.classFee) || (periodsCount * rate);
    var mentorFee = Number(updatedData.mentorFee) || 0;
    var totalFee = Number(updatedData.totalFee) || (classFee + mentorFee);
    var payMode = updatedData.payMode || 'perPeriod';
    var payNote = updatedData.payNote || '';
    var isActingMentor = (payMode === 'mentorDaily' || mentorFee > 0);

    // 更新總表
    if (date) summarySheet.getRange(targetRowIndex, 3).setValue(date);
    summarySheet.getRange(targetRowIndex, 6).setValue(absentTeacher);
    summarySheet.getRange(targetRowIndex, 7).setValue(subTeacher);
    summarySheet.getRange(targetRowIndex, 8).setValue(className);
    summarySheet.getRange(targetRowIndex, 9).setValue(leaveType);
    summarySheet.getRange(targetRowIndex, 10).setValue(periodsCount);
    summarySheet.getRange(targetRowIndex, 11).setValue(rate);
    summarySheet.getRange(targetRowIndex, 12).setValue(classFee);
    summarySheet.getRange(targetRowIndex, 13).setValue(isActingMentor ? "是" : "否");
    summarySheet.getRange(targetRowIndex, 15).setValue(mentorFee);
    summarySheet.getRange(targetRowIndex, 16).setValue(totalFee);
    summarySheet.getRange(targetRowIndex, 17).setValue(payMode);
    summarySheet.getRange(targetRowIndex, 18).setValue(payNote);

    // 同步更新明細表
    if (detailSheet) {
      var detLastRow = detailSheet.getLastRow();
      if (detLastRow > 1) {
        var detValues = detailSheet.getRange(2, 1, detLastRow - 1, DETAIL_HEADERS.length).getValues();
        for (var j = 0; j < detValues.length; j++) {
          if (String(detValues[j][0] || '').trim() === String(appId).trim()) {
            var detRowNum = j + 2;
            if (date) detailSheet.getRange(detRowNum, 2).setValue(date);
            if (className) detailSheet.getRange(detRowNum, 3).setValue(className);
            if (subTeacher) detailSheet.getRange(detRowNum, 6).setValue(subTeacher);
          }
        }
      }
    }

    return {
      status: "success",
      message: "🎉 代課單 " + appId + " 已成功更新！",
      records: buildRecordsForMonth(ss, targetMonth)
    };
  } catch (err) {
    return { status: "error", message: "更新失敗：" + err.toString() };
  }
}

function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    return dateVal.getFullYear() + "-" +
      ("0" + (dateVal.getMonth() + 1)).slice(-2) + "-" +
      ("0" + dateVal.getDate()).slice(-2);
  }
  var str = dateVal.toString();
  if (str.indexOf("GMT") !== -1 || str.indexOf("T") !== -1) {
    var dObj = new Date(str);
    return dObj.getFullYear() + "-" +
      ("0" + (dObj.getMonth() + 1)).slice(-2) + "-" +
      ("0" + dObj.getDate()).slice(-2);
  }
  return str;
}
