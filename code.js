/************************************************************
 * 🚢 Shipping Schedule Web App (입력 + 수정) — 로그 기반 최종본
 * - DB 시트: "로그" (A~AB, 28개 컬럼)
 * - 엑셀 템플릿: "항목명" 시트 1행(A~U) 그대로 사용 (21개)
 * - ✅ 수정 제출 시에도 신규와 동일하게 엑셀 생성하여 첨부
 ************************************************************/

var LOG_SHEET_NAME = "로그";
var HEADER_SHEET_NAME = "항목명";

/************************************************************
 * 로그 시트 헤더 (A~AB)
 ************************************************************/
var LOG_HEADERS = [
  "구분",               // A
  "입력시간",           // B
  "발신자 이메일",      // C
  "이메일 제목",        // D
  "L/C번호(30)",        // E
  "포워더(10)",         // F
  "컨테이너대수(100)",  // G
  "BULK/CNTR 구분(100)",// H
  "VESSEL & VOY(100)",  // I
  "서류마감일(100)",    // J
  "CARGO 마감일(8)",    // K
  "출항장소(100)",      // L
  "출항일(ETD)(8)",     // M
  "PORT명(100)",        // N
  "도착일(ETA)(8)",     // O
  "선사명(LINE)(100)",  // P
  "CFS / CY(100)",      // Q
  "담당자(100)",        // R
  "담당자 연락처(100)", // S
  "BOOKING NO(100)",    // T
  "장지장코드(숫자)(100)", // U
  "CFS/CY주소(100)",    // V
  "CFS(CY)코드",        // W  ← 이게 꼭 들어가야 함
  "CALL SIGN",          // X
  "항구청코드",         // Y
  "담당자 이메일",      // Z
  "추가 CC",            // AA
  "비고"                // AB
];

/************************************************************
 * LOG_HEADERS 내에서 특정 헤더명의 인덱스를 반환
 ************************************************************/
function idxInLog_(headerName) {
  for (var i = 0; i < LOG_HEADERS.length; i++) {
    if (LOG_HEADERS[i] === headerName) return i;
  }
  return -1;
}

/************************************************************
 * WebApp 진입점
 *  - /exec           → 입력 화면
 *  - /exec?mode=edit → 수정 화면
 ************************************************************/
function doGet(e) {
  var mode = e && e.parameter && e.parameter.mode;
  if (mode === "edit") {
    return HtmlService.createTemplateFromFile("index_edit")
      .evaluate()
      .setTitle("Shipping Schedule 수정")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Shipping Schedule 입력")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/************************************************************
 * HTML include
 ************************************************************/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/************************************************************
 * 로그 시트 핸들러
 ************************************************************/
function getLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(LOG_SHEET_NAME);

  // 1행 헤더 강제 세팅
  var firstRow = sheet.getRange(1, 1, 1, LOG_HEADERS.length).getValues()[0];
  var needUpdate = false;

  for (var i = 0; i < LOG_HEADERS.length; i++) {
    if (firstRow[i] !== LOG_HEADERS[i]) { needUpdate = true; break; }
  }
  if (needUpdate) {
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  }
  return sheet;
}

/************************************************************
 * 항목명/코드명/담당자 불러오기 (입력·수정 공통)
 ************************************************************/
function loadFormData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 항목명: 엑셀 템플릿용 헤더 (A1:U1, 21개)
  var sheetA = ss.getSheetByName(HEADER_SHEET_NAME);
  var headers = sheetA.getRange(1, 1, 1, 21).getValues()[0];

  // 코드명: select 용 코드/상세
  var sheetC = ss.getSheetByName("코드명");
  var lastRowC = sheetC.getLastRow();
  var codeMap = {};
  if (lastRowC > 1) {
    var dataC = sheetC.getRange(2, 1, lastRowC - 1, 3).getValues();
    for (var i = 0; i < dataC.length; i++) {
      var r = dataC[i];
      var g = r[0], code = r[1], detail = r[2];
      if (!g) continue;
      if (!codeMap[g]) codeMap[g] = [];
      codeMap[g].push({ code: code, detail: detail });
    }
  }

  // 담당자: 해외영업팀만 사용
  var sheetD = ss.getSheetByName("담당자");
  var lastRowD = sheetD.getLastRow();
  var managers = [];
  if (lastRowD > 1) {
    var dataD = sheetD.getRange(2, 1, lastRowD - 1, 3).getValues();
    for (var j = 0; j < dataD.length; j++) {
      var rr = dataD[j];
      if (rr[0] === "해외영업팀") {
        managers.push({ team: rr[0], name: rr[1], email: rr[2] });
      }
    }
  }

  return { headers: headers, codeMap: codeMap, managers: managers };
}

/************************************************************
 * 물류팀 CC 이메일 로드
 ************************************************************/
function getLogisticsEmails_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("담당자");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (r[0] === "물류팀" && r[2]) list.push(String(r[2]).trim());
  }
  return list;
}

/************************************************************
 * CC 문자열 → 유효 이메일만 추출
 ************************************************************/
function parseExtraEmails_(raw) {
  if (!raw) return [];
  var regex = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
  var parts = String(raw).split(",");
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var e = parts[i].trim();
    if (e && regex.test(e)) out.push(e);
  }
  return out;
}

/************************************************************
 * HTML escape
 ************************************************************/
function escapeHtml_(s) {
  if (s === null || s === undefined) return "";
  return String(s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/************************************************************
 * 신규 제출 요약표 (이메일용)
 ************************************************************/
function buildSummaryHtml_(headers, values, note) {
  var html = '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:13px;">';
  for (var i = 0; i < headers.length; i++) {
    html += '<tr>' +
      '<td style="font-weight:bold;background:#f3f3f3;">' + escapeHtml_(headers[i]) + '</td>' +
      '<td>' + escapeHtml_(values[i]) + '</td>' +
      '</tr>';
  }
  if (note) {
    html += '<tr><td style="font-weight:bold;background:#f3f3f3;">비고</td><td>' + escapeHtml_(note) + '</td></tr>';
  }
  html += '</table>';
  return html;
}

/************************************************************
 * ✅ 엑셀 생성 공통 함수 (신규/수정 공용)
 * - logRow : LOG_HEADERS 기준 1행 데이터
 * - 반환   : XLSX Blob
 ************************************************************/
function buildShippingExcelBlob_(logRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headerSheet = ss.getSheetByName(HEADER_SHEET_NAME);
  var headerRow = headerSheet.getRange(1, 1, 1, headerSheet.getLastColumn()).getValues()[0];

  // 엑셀용 값: 항목명 순서대로 매핑
  var excelValues = [];
  for (var i = 0; i < headerRow.length; i++) {
    var li = idxInLog_(headerRow[i]);
    excelValues.push(li >= 0 ? logRow[li] : "");
  }

  // 임시 스프레드시트 생성
  var tmp = SpreadsheetApp.create("export_temp");
  var ts = tmp.getSheets()[0];
  ts.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  ts.getRange(2, 1, 1, headerRow.length).setNumberFormat("@");   // TEXT 강제
  ts.getRange(2, 1, 1, headerRow.length).setValues([excelValues]);
  SpreadsheetApp.flush();

  // XLSX 변환
  var resp = UrlFetchApp.fetch(
    "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + tmp.getId() + "&exportFormat=xlsx",
    { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } }
  );

  var xlsxBlob = resp.getBlob().setName("SHIPPING SCHEDULE.xlsx");

  // 임시 파일 삭제
  DriveApp.getFileById(tmp.getId()).setTrashed(true);

  return xlsxBlob;
}

/************************************************************
 * values[] + 메타 → 로그 1행 데이터 생성
 * - typeFlag: "신규" / "수정"
 * - values : 항목명 21개 + 담당자이메일(맨 뒤)
 ************************************************************/
function buildLogRowFromValues_(typeFlag, values, extraCcRaw, note, fileData, senderEmail, subject) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headerSheet = ss.getSheetByName(HEADER_SHEET_NAME);
  var headerRow = headerSheet.getRange(1, 1, 1, headerSheet.getLastColumn()).getValues()[0]; // 21개

  var rowData = [];
  for (var i = 0; i < LOG_HEADERS.length; i++) rowData.push("");

  var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var finalSubject = subject && subject !== "" ? subject : (typeFlag === "신규" ? "SHIPPING SCHEDULE" : "[수정건] SHIPPING SCHEDULE");
  var sender = senderEmail || "";

  var fileNames = [];
  if (fileData && fileData.length) {
    for (var f = 0; f < fileData.length; f++) {
      var fd = fileData[f];
      if (fd && fd.name) fileNames.push(fd.name);
    }
  }

  var logNote = note || "";
  if (fileNames.length) {
    if (logNote) logNote += "\n첨부: " + fileNames.join(", ");
    else logNote = "첨부: " + fileNames.join(", ");
  }

  rowData[idxInLog_("구분")] = typeFlag;
  rowData[idxInLog_("입력시간")] = now;
  rowData[idxInLog_("발신자 이메일")] = sender;
  rowData[idxInLog_("이메일 제목")] = finalSubject;

  // 항목명(21개) 매핑
  for (var i2 = 0; i2 < headerRow.length; i2++) {
    var hName = headerRow[i2];
    var li = idxInLog_(hName);
    if (li >= 0 && i2 < values.length) rowData[li] = values[i2] || "";
  }

  // 담당자 이메일 (values 마지막)
  var managerEmail = "";
  if (values.length > headerRow.length) managerEmail = values[headerRow.length];
  rowData[idxInLog_("담당자 이메일")] = managerEmail || "";

  rowData[idxInLog_("추가 CC")] = extraCcRaw || "";
  rowData[idxInLog_("비고")] = logNote;

  return {
    rowData: rowData,
    managerEmail: managerEmail,
    subject: finalSubject,
    fileNames: fileNames
  };
}

/************************************************************
 * 신규 제출 (입력 화면) — 실제 저장 + 메일 발송
 ************************************************************/
function submitData(values, extra, note, fileData, senderEmail, subject) {
  var logSheet = getLogSheet_();

  var built = buildLogRowFromValues_(
    "신규",
    values,
    extra,
    note || "",
    fileData,
    senderEmail,
    subject
  );

  var destRow = logSheet.getLastRow() + 1;
  logSheet.getRange(destRow, 1, 1, LOG_HEADERS.length)
    .setValues([built.rowData]);

  // 메일 발송 (엑셀 첨부)
  sendNewMail_(destRow, built, extra, note || "", fileData, senderEmail);

  return true;
}

/************************************************************
 * 신규 제출 메일 발송 (엑셀 첨부)
 ************************************************************/
function sendNewMail_(rowNum, built, extra, note, fileData, senderEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headerSheet = ss.getSheetByName(HEADER_SHEET_NAME);
  var headerRow = headerSheet.getRange(1, 1, 1, headerSheet.getLastColumn()).getValues()[0];

  var logSheet = getLogSheet_();
  var row = logSheet.getRange(rowNum, 1, 1, LOG_HEADERS.length).getValues()[0];

  // ✅ 신규: 엑셀 생성
  var xlsxBlob = buildShippingExcelBlob_(row);

  // TO
  var toEmail = row[idxInLog_("담당자 이메일")];

  // CC 구성
  var ccList = [];
  var lg = getLogisticsEmails_();
  for (var i2 = 0; i2 < lg.length; i2++) ccList.push(lg[i2]);

  var extraList = parseExtraEmails_(extra);
  for (var j = 0; j < extraList.length; j++) ccList.push(extraList[j]);

  var uniqueCc = {};
  var ccFinal = [];
  for (var c = 0; c < ccList.length; c++) {
    var e = ccList[c];
    if (e && e !== toEmail && !uniqueCc[e]) {
      uniqueCc[e] = true;
      ccFinal.push(e);
    }
  }

  // 첨부파일: 엑셀 + 추가 업로드 파일
  var attachments = [xlsxBlob];
  if (fileData && fileData.length) {
    for (var k = 0; k < fileData.length; k++) {
      var f = fileData[k];
      attachments.push(
        Utilities.newBlob(
          Utilities.base64Decode(f.data),
          f.type,
          f.name
        )
      );
    }
  }

  // 본문 요약(템플릿 항목 순서 기준으로 표 생성)
  var excelValues = [];
  for (var i = 0; i < headerRow.length; i++) {
    var li = idxInLog_(headerRow[i]);
    excelValues.push(li >= 0 ? String(row[li]) : "");
  }

  var html = buildSummaryHtml_(headerRow, excelValues, note);

  var options = {
    htmlBody: "포워더가 제출한 Shipping Schedule 입니다.<br><br>" + html,
    attachments: attachments
  };
  if (ccFinal.length) options.cc = ccFinal.join(",");
  if (senderEmail) options.replyTo = senderEmail;

  MailApp.sendEmail(toEmail, built.subject, "", options);
}

/************************************************************
 * 🔍 최신 제출분 조회 (BOOKING NO(100) 기준, 가장 최근 행 반환)
 ************************************************************/
function findLatestRecord(bookingNo) {
  bookingNo = String(bookingNo || "").trim();
  if (!bookingNo) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = getLogSheet_();
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) return null;

  var bookIdx = idxInLog_("BOOKING NO(100)");
  if (bookIdx < 0) return null;

  var data = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  var targetRow = -1;

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][bookIdx]).trim() === bookingNo) targetRow = i + 2;
  }
  if (targetRow === -1) return null;

  var row = logSheet.getRange(targetRow, 1, 1, LOG_HEADERS.length).getValues()[0];

  // 항목명 헤더
  var headerSheet = ss.getSheetByName(HEADER_SHEET_NAME);
  var headers = headerSheet.getRange(1, 1, 1, headerSheet.getLastColumn()).getValues()[0];

  // 항목명 순서대로 값 복원
  var values21 = [];
  for (var h = 0; h < headers.length; h++) {
    var hName = headers[h];
    var li = idxInLog_(hName);
    values21.push(li >= 0 ? row[li] : "");
  }

  return {
    row: targetRow,
    headers: headers,
    values: values21,
    managerEmail: row[idxInLog_("담당자 이메일")] || "",
    senderEmail: row[idxInLog_("발신자 이메일")] || "",
    subject: row[idxInLog_("이메일 제목")] || "",
    extraCc: row[idxInLog_("추가 CC")] || "",
    note: row[idxInLog_("비고")] || ""
  };
}

/************************************************************
 * ✏ 수정 제출 (로그에 새 "수정" 행 + 변경항목 빨간색 + 메일 + ✅엑셀첨부)
 ************************************************************/
function submitEdit(bookingNo, values, extra, note, fileData, senderEmail, subject) {
  bookingNo = String(bookingNo || "").trim();
  if (!bookingNo) return "NOT_FOUND";

  var logSheet = getLogSheet_();
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) return "NOT_FOUND";

  var bookIdx = idxInLog_("BOOKING NO(100)");
  if (bookIdx < 0) return "NOT_FOUND";

  var data = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  var targetRow = -1;

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][bookIdx]).trim() === bookingNo) targetRow = i + 2;
  }
  if (targetRow === -1) return "NOT_FOUND";

  var oldRow = logSheet.getRange(targetRow, 1, 1, LOG_HEADERS.length).getValues()[0];

  // 새 행 데이터 생성
  var built = buildLogRowFromValues_("수정", values, extra, note, fileData, senderEmail, subject);
  var newRow = built.rowData;

  // 변경 여부 체크 (구분/입력시간 제외)
  var colors = [];
  var idxFlag = idxInLog_("구분");
  var idxTime = idxInLog_("입력시간");

  for (var c = 0; c < LOG_HEADERS.length; c++) {
    var before = String(oldRow[c] || "");
    var after = String(newRow[c] || "");
    var changed = (before !== after);

    if (c === idxFlag || c === idxTime) changed = false;
    colors.push(changed ? "#d00000" : "#000000");
  }

  var destRow = logSheet.getLastRow() + 1;
  logSheet.getRange(destRow, 1, 1, LOG_HEADERS.length).setValues([newRow]);
  logSheet.getRange(destRow, 1, 1, LOG_HEADERS.length).setFontColors([colors]);

  // ✅ 수정 메일 발송 (엑셀 포함)
  sendEditMail_(oldRow, newRow, built, extra, note, fileData, senderEmail);

  return "OK";
}

/************************************************************
 * 수정 메일 발송 (변경항목 하이라이트) + ✅엑셀 생성 첨부
 ************************************************************/
function sendEditMail_(oldRow, newRow, built, extra, note, fileData, senderEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headerSheet = ss.getSheetByName(HEADER_SHEET_NAME);
  var headerRow = headerSheet.getRange(1, 1, 1, headerSheet.getLastColumn()).getValues()[0];

  var toEmail = newRow[idxInLog_("담당자 이메일")];

  // CC 구성
  var ccList = [];
  var lg = getLogisticsEmails_();
  for (var i = 0; i < lg.length; i++) ccList.push(lg[i]);

  var extraList = parseExtraEmails_(extra);
  for (var j = 0; j < extraList.length; j++) ccList.push(extraList[j]);

  var uniqueCc = {};
  var ccFinal = [];
  for (var k = 0; k < ccList.length; k++) {
    var e = ccList[k];
    if (e && e !== toEmail && !uniqueCc[e]) {
      uniqueCc[e] = true;
      ccFinal.push(e);
    }
  }

  // ✅ 수정: 엑셀 생성 + 추가 업로드 파일
  var xlsxBlob = buildShippingExcelBlob_(newRow);

  var attachments = [xlsxBlob];
  if (fileData && fileData.length) {
    for (var a = 0; a < fileData.length; a++) {
      var f = fileData[a];
      attachments.push(
        Utilities.newBlob(
          Utilities.base64Decode(f.data),
          f.type,
          f.name
        )
      );
    }
  }

  // 본문(변경 항목 강조)
  var body = "📌 Shipping Schedule 수정 안내<br><br>";
  var booking = newRow[idxInLog_("BOOKING NO(100)")] || "";
  body += "<b>BOOKING NO(100) :</b> " + escapeHtml_(String(booking)) + "<br><br>";
  body += '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:13px;">';

  // 메일 표: 항목명(21개) + 담당자 이메일
  var namesForMail = [];
  for (var h = 0; h < headerRow.length; h++) namesForMail.push(headerRow[h]);
  namesForMail.push("담당자 이메일");

  for (var r = 0; r < namesForMail.length; r++) {
    var hName = namesForMail[r];
    var li = idxInLog_(hName);

    var before = (li >= 0) ? String(oldRow[li] || "") : "";
    var after = (li >= 0) ? String(newRow[li] || "") : "";
    var changed = (before !== after);

    body += '<tr' + (changed ? ' style="background:#fff2cc;"' : '') + '>';
    body += '<td style="font-weight:bold;background:#f3f3f3;">' + escapeHtml_(hName) + '</td>';

    if (changed) {
      body += '<td><b>' + escapeHtml_(after) + '</b> <span style="color:#d00000">(기존: ' + escapeHtml_(before) + ')</span></td>';
    } else {
      body += '<td>' + escapeHtml_(after) + '</td>';
    }
    body += '</tr>';
  }

  body += "</table>";

  if (note) {
    body += "<br><b>비고:</b><br>" + escapeHtml_(note);
  }

  var options = { htmlBody: body };
  if (attachments.length) options.attachments = attachments;
  if (ccFinal.length) options.cc = ccFinal.join(",");
  if (senderEmail) options.replyTo = senderEmail;

  MailApp.sendEmail(toEmail, built.subject, "", options);
}
