/**
 * 이정기 모바일 명함 — 방명록 백엔드 (Google Apps Script)
 *
 * 시트 구성(첫 줄 머리글): 일시 | 이름 | 메시지 | 승인
 *  - 새 글은 '승인' 칸이 체크 해제(FALSE) 상태로 쌓입니다.
 *  - 게시하려면 '승인' 칸을 체크하면 됩니다. 지우려면 행을 삭제하면 됩니다.
 */

const SHEET_NAME = '방명록';
const MAX_NAME = 20;
const MAX_MESSAGE = 300;
const SHOW_LIMIT = 50;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['일시', '이름', '메시지', '승인']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 승인된 글 목록을 최신순으로 돌려줍니다. */
function doGet() {
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues().slice(1);
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  const items = rows
    .filter(r => r[3] === true)
    .reverse()
    .slice(0, SHOW_LIMIT)
    .map(r => ({
      date: Utilities.formatDate(new Date(r[0]), tz, 'yyyy.MM.dd'),
      name: String(r[1]),
      message: String(r[2])
    }));
  return json_({ ok: true, items: items });
}

/** 새 글을 '미승인' 상태로 저장합니다. */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.website) return json_({ ok: true }); // 자동등록 방지 함정 칸
    const name = String(data.name || '').trim().slice(0, MAX_NAME);
    const message = String(data.message || '').trim().slice(0, MAX_MESSAGE);
    if (!name || message.length < 2) {
      return json_({ ok: false, error: '이름과 말씀을 확인해 주세요.' });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const sh = getSheet_();
      const row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 4).setValues([[new Date(), name, message, false]]);
      sh.getRange(row, 4).insertCheckboxes();
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: '저장하지 못했습니다. 잠시 뒤 다시 시도해 주세요.' });
  }
}
