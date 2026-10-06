/**
 * 아트에이블 수채화다이어리 — 뉴스레터 신청 → 구글 시트 저장
 *
 * 설치 방법
 * 1) 새 구글 스프레드시트를 만든다 (예: "수채화다이어리 뉴스레터")
 * 2) 메뉴 [확장 프로그램] > [Apps Script] 에 이 코드를 통째로 붙여넣고 저장
 * 3) [배포] > [새 배포] > 유형 '웹 앱'
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 * 4) 권한 승인 후 나오는 웹 앱 URL(https://script.google.com/macros/s/.../exec)을
 *    index.html 의 const ENDPOINT='' 에 넣는다
 */

const SHEET_NAME = '신청자';
const HEADERS = ['신청일시', '이름', '이메일', '개인정보 동의', '유입 페이지'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const email = String(p.email || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json({ ok: false, error: 'invalid email' });

    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last > 1) {
      const emails = sheet.getRange(2, 3, last - 1, 1).getValues().flat().map(String);
      if (emails.includes(email)) return json({ ok: true, duplicate: true });
    }

    sheet.appendRow([
      new Date(),
      String(p.name || '').trim(),
      email,
      p.consent ? '동의' : '',
      String(p.source || '')
    ]);
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true, message: 'newsletter endpoint alive' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sheet;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
