/**
 * 청첩장 모임 RSVP -> 구글 시트 저장용 Apps Script
 * 사용법은 README-응답연동.md 참고
 */
const SPREADSHEET_ID = "";   // 시트 주소 .../d/여기부분/edit 의 ID. 시트에서 확장프로그램>Apps Script로 열었다면 비워둬도 됨
const SHEET_NAME = "응답";   // 요청에 sheet 값이 없을 때 쓰는 기본 탭 이름
const HEADERS = ["제출시각", "이름", "참석여부", "인원", "동행인", "메모"];

function getSS_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("스프레드시트를 찾을 수 없습니다. SPREADSHEET_ID를 입력해주세요.");
  return ss;
}

function save_(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getSS_();
    const tab = String(p.sheet || SHEET_NAME).replace(/[\[\]\*\?\/\\:]/g, "").slice(0, 90) || SHEET_NAME;
    const sh = ss.getSheetByName(tab) || ss.insertSheet(tab);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#dbe8f4");
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), p.name || "", p.attend || "", Number(p.count) || 0, p.mates || "", p.note || ""]);
    sh.getRange("H1").setValue("참석 인원 합계");
    sh.getRange("I1").setFormula('=SUMIF(C:C,"참석",D:D)');
    sh.getRange("H2").setValue("참석 응답 수");
    sh.getRange("I2").setFormula('=COUNTIF(C:C,"참석")');
    sh.getRange("H3").setValue("불참 응답 수");
    sh.getRange("I3").setFormula('=COUNTIF(C:C,"불참")');
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  try {
    save_(e.parameter || {});
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err.message);
  }
}

function doGet() {
  return ContentService.createTextOutput("RSVP endpoint is running");
}

/** 편집기에서 이 함수를 선택해 '실행'하면 테스트 행이 들어갑니다 (권한 승인도 이때 진행) */
function testRow() {
  save_({ name: "테스트", attend: "참석", count: 2, mates: "동행테스트", note: "삭제해도 됨" });
}
