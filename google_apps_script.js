/**
 * 代課費計算與管理系統 - Google Apps Script (GAS) 雲端資料庫後台
 * 【按月分表儲存 (總表 + 明細表)】
 *
 * 說明：
 * 1. 依代課日期的月份，自動建立 `YYYY-MM_總表` 與 `YYYY-MM_明細` 兩個分頁。
 * 2. 歷史月份各自獨立，不互相干擾。
 * 3. 行政端帶 `?month=YYYY-MM` 即可精準取得該月資料。
 *
 * ⚠️ 欄位相容性（勿隨意刪欄）：
 *   admin.html 需要 subTeacher / periodsCount / className / classFee / mentorFee / totalFee
 *   generate_voucher.py 另需 rate / payMode / actingDays / isActingMentor / periodsDetail[].paid
 *   總表與明細的欄位順序由下方 SUMMARY_HEADERS / DETAIL_HEADERS 定義，讀寫共用同一份定義。
 */

/**
 * 讀取金鑰（共用密碼）
 *
 * 用途：teacher.html 是公開網頁，網址會寫在原始碼裡任何人都看得到。
 * 若後端不設限，任何人都能用 `?month=all` 抓走全部代課費資料（教師姓名、金額）。
 * 因此「讀取」一律要求帶上金鑰，只有 admin.html（僅在承辦人電腦上）知道。
 *
 * ⚠️ 注意事項：
 * - 「寫入」(doPost) 維持開放，老師才能免登入填報。
 * - 課表 (?action=timetable) 維持開放，teacher.html 需要它顯示課程。
 * - 金鑰外流時，改掉這一行後重新部署即可（網址不變）。
 * - admin.html 不會寫死金鑰，由承辦人在齒輪設定中填入、存於瀏覽器。
 */
var SECRET_KEY = '請在此填入自訂的隨機安全金鑰密碼';;

// 課表分頁名稱（每學年更新課表時，直接覆蓋這個分頁的內容即可，不必改程式碼）
var TIMETABLE_SHEET = 'Class_Timetables';

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

function getCorsResponse(content, type) {
  var output = ContentService.createTextOutput(content);
  if (type === 'json') {
    output.setMimeType(ContentService.MimeType.JSON);
  } else {
    output.setMimeType(ContentService.MimeType.TEXT);
  }
  return output;
}

// 取得（或建立）指定月份的分頁，並確保標頭列存在
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 接收教師端網頁送出的資料 (POST)
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var dateStr = data.date; // 格式為 YYYY-MM-DD
    if (!dateStr) {
      return getCorsResponse(JSON.stringify({ status: "error", message: "缺少代課日期 (date)！" }), 'json');
    }

    var month = dateStr.substring(0, 7); // YYYY-MM
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = getOrCreateSheet(ss, month + "_總表", SUMMARY_HEADERS);
    var detailSheet = getOrCreateSheet(ss, month + "_明細", DETAIL_HEADERS);

    var rate = Number(data.rate) || 0;
    var payMode = data.payMode || 'perPeriod';
    // 整天代理導師模式：金額由行政人員自行計算，此處不臆測費用
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

    // 明細：一節課一列
    if (data.periodsDetail && Array.isArray(data.periodsDetail)) {
      for (var i = 0; i < data.periodsDetail.length; i++) {
        var p = data.periodsDetail[i];
        var isPaid = p.paid === true;
        // 未計費的節次（午餐、午休、放學等指導）費率與金額皆為 0
        var periodRate = isPaid ? rate : 0;

        detailSheet.appendRow([
          appId,
          data.date,
          data.className || "",
          p.period,
          p.subject || "",
          p.subTeacher || data.subTeacher || "",
          isPaid ? "是" : "否",
          periodRate,
          periodRate,
          "待核對", // 預設狀態
          p.handover || "",
          "" // 確認日期初始為空
        ]);
      }
    }

    return getCorsResponse(JSON.stringify({ status: "success", appId: appId, message: "資料已成功按月寫入 Google 試算表！" }), 'json');
  } catch (error) {
    return getCorsResponse(JSON.stringify({ status: "error", message: error.toString() }), 'json');
  }
}

/**
 * 讀取 Class_Timetables 分頁，組成 teacher.html 可直接使用的巢狀結構：
 *   { "四甲": { 1: { 1: { subject: "數學", teacher: "楊琬瑄" } } } }
 * 分頁欄位：A=班級 B=星期(1-5) C=節次 D=科目 E=授課教師（第 1 列為標頭）
 */
function buildTimetable(ss) {
  var sheet = ss.getSheetByName(TIMETABLE_SHEET);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var db = {};
  for (var i = 0; i < values.length; i++) {
    var cls = String(values[i][0] || '').trim();
    var day = parseInt(values[i][1], 10);
    var period = parseInt(values[i][2], 10);
    var subject = String(values[i][3] || '').trim();
    var teacher = String(values[i][4] || '').trim();
    // 跳過空白列或格式不符的列，避免課表壞掉
    if (!cls || !day || !period) continue;

    if (!db[cls]) db[cls] = {};
    if (!db[cls][day]) db[cls][day] = {};
    db[cls][day][period] = { subject: subject, teacher: teacher };
  }
  return db;
}

// 列出試算表中所有有資料的月份（依 `YYYY-MM_總表` 分頁名稱推斷），由舊到新排序
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

// 讀取單一月份，組裝成主從結構 (Master-Detail)。找不到分頁或無資料時回傳空陣列。
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

  // 明細依 appId 分組
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
      periodsCount: Number(s[9]) || 0,
      rate: Number(s[10]) || 0,
      classFee: Number(s[11]) || 0,
      isActingMentor: s[12] === "是",
      actingDays: Number(s[13]) || 0,
      mentorFee: Number(s[14]) || 0,
      totalFee: Number(s[15]) || 0,
      payMode: s[16] || 'perPeriod',
      payNote: s[17] || "",
      periodsDetail: detailsMap[currentAppId] || []
    });
  }
  return records;
}

/**
 * 提供行政端網頁讀取與修改資料 (GET)
 *
 * 【公開】不需金鑰：
 * - 取得課表：GET ?action=timetable        （teacher.html 使用）
 *
 * 【受保護】需帶 &key=<金鑰>：
 * - 讀取某月：GET ?month=YYYY-MM&key=...
 * - 讀取全部：GET ?month=all&key=...        （admin.html「備份所有 JSON」）
 * - 修改狀態：GET ?action=confirm&month=YYYY-MM&appId=SUB-XXX&period=節次&key=...
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var params = e.parameter;

    // --- 課表（公開，不需金鑰）---
    // teacher.html 為公開網頁，無法安全保存金鑰；課表本身不含金額資訊。
    if (params.action === "timetable") {
      return getCorsResponse(JSON.stringify(buildTimetable(ss)), 'json');
    }

    // --- 以下皆需金鑰 ---
    // 沒有這道檢查，任何人拿到網址就能用 ?month=all 抓走全部代課費資料。
    if (params.key !== SECRET_KEY) {
      return getCorsResponse(JSON.stringify({
        status: "error",
        message: "未授權：缺少或錯誤的金鑰。請在 admin.html 右上角齒輪設定中填入正確金鑰。"
      }), 'json');
    }

    // --- 狀態修改 (action = confirm) ---
    if (params.action === "confirm" && params.month && params.appId && params.period) {
      var detailSheet = ss.getSheetByName(params.month + "_明細");
      if (!detailSheet) {
        return getCorsResponse(JSON.stringify({ status: "error", message: "找不到該月份的明細分頁！" }), 'json');
      }

      var lastRow = detailSheet.getLastRow();
      if (lastRow <= 1) {
        return getCorsResponse(JSON.stringify({ status: "error", message: "該月份明細分頁尚無資料！" }), 'json');
      }

      var values = detailSheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues();
      var success = false;

      for (var i = 0; i < values.length; i++) {
        // 比對 申請編號(ID) 與 節次(Period)
        if (values[i][0] === params.appId && values[i][3].toString() === params.period.toString()) {
          var rowNum = i + 2; // 還原為試算表列號（1-indexed 且跳過標頭）
          detailSheet.getRange(rowNum, DETAIL_COL_STATUS).setValue("已確認");
          detailSheet.getRange(rowNum, DETAIL_COL_CONFIRM_DATE).setValue(new Date().toISOString());
          success = true;
          break;
        }
      }

      return getCorsResponse(JSON.stringify(success
        ? { status: "success", message: "狀態已更新為已確認！" }
        : { status: "error", message: "找不到對應的代課節次明細！" }), 'json');
    }

    // --- 讀取全部月份 (month=all) ---
    // 供 admin.html「備份所有 JSON」使用。按月分表後，單月請求只會拿到一個月，
    // 若沒有這個分支，「全部備份」會變成只備份單月的假備份。
    if (params.month === "all") {
      var months = listDataMonths(ss);
      var allRecords = [];
      for (var n = 0; n < months.length; n++) {
        allRecords = allRecords.concat(buildRecordsForMonth(ss, months[n]));
      }
      allRecords.sort(function (a, b) {
        return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0);
      });
      return getCorsResponse(JSON.stringify(allRecords), 'json');
    }

    // --- 讀取單一月份 ---
    var targetMonth = params.month;
    if (!targetMonth) {
      var today = new Date();
      targetMonth = today.getFullYear() + "-" + ("0" + (today.getMonth() + 1)).slice(-2);
    }

    return getCorsResponse(JSON.stringify(buildRecordsForMonth(ss, targetMonth)), 'json');
  } catch (error) {
    return getCorsResponse(JSON.stringify({ status: "error", message: error.toString() }), 'json');
  }
}

// 輔助函式：確保日期格式為 YYYY-MM-DD
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
