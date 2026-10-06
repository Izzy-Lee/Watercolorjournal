/**
 * 아트에이블 수채화다이어리 — 뉴스레터 신청 → 구글 시트 저장 (v2)
 *
 * 시트에서 [확장 프로그램 > Apps Script]로 만든 경우: 그 시트에 저장
 * script.google.com에서 따로 만든 경우: 저장용 시트를 자동으로 만들어 그곳에 저장
 *   → 웹 앱 주소를 브라우저로 열면 저장 중인 시트 주소(sheetUrl)가 보입니다.
 *
 * 코드 수정 후에는 [배포 > 배포 관리 > 연필 > 버전: 새 버전 > 배포] (주소 유지)
 */

const SHEET_NAME = '신청자';
const HEADERS = ['신청일시', '이름', '이메일', '개인정보 동의', '유입 페이지'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = params_(e);
    const email = String(p.email || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      console.warn('invalid email', JSON.stringify(p));
      return json({ ok: false, error: 'invalid email' });
    }

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
  } catch (err) {
    console.error(err && err.stack || err);
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  try {
    return json({ ok: true, message: 'newsletter endpoint alive', sheetUrl: getSpreadsheet_().getUrl() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// e.parameter가 비어 있으면 본문을 직접 해석
function params_(e) {
  const p = Object.assign({}, (e && e.parameter) || {});
  if (!p.email && e && e.postData && e.postData.contents) {
    const body = e.postData.contents;
    try {
      Object.assign(p, JSON.parse(body));
    } catch (_) {
      body.split('&').forEach(function (kv) {
        const i = kv.indexOf('=');
        if (i < 0) return;
        const k = decodeURIComponent(kv.slice(0, i).replace(/\+/g, ' '));
        p[k] = decodeURIComponent(kv.slice(i + 1).replace(/\+/g, ' '));
      });
    }
  }
  return p;
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('수채화다이어리 뉴스레터 신청자');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getSheet_() {
  const ss = getSpreadsheet_();
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
