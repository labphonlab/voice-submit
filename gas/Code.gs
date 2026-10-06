/**
 * 音声課題の受付（Google Apps Script・スプレッドシートに紐づけて使う）
 *
 * シート「課題」     … 先生が課題を書く（1行＝1項目）
 * シート「受信ファイル」… 届いた録音1件ごとの記録（自動）
 * シート「提出一覧」   … 学生1人の1回の提出ごとの記録（自動）
 * 録音はドライブの「音声課題_提出/<課題ID>/<学籍番号>_<氏名>/」に WAV で保存する。
 */

const SHEET_TASKS = '課題';
const SHEET_FILES = '受信ファイル';
const SHEET_SUBS = '提出一覧';
const ROOT_FOLDER_NAME = '音声課題_提出';
const MAX_BYTES = 25 * 1024 * 1024; // 1項目あたりの上限
const TZ = 'Asia/Tokyo';

const TASK_HEADER = ['課題ID', '課題名', '説明', '締切', '項目番号', '提示文', '読み・補足', '最長秒数'];
const FILES_HEADER = ['受信日時', '提出ID', '課題ID', '学籍番号', '氏名', '項目番号', '提示文', '秒数', 'サンプリング周波数', 'ピーク', 'ファイル名', 'URL'];
const SUBS_HEADER = ['完了日時', '受付番号', '課題ID', '学籍番号', '氏名', '項目数', '受信数', '状態', '提出ID', '端末'];

/** 最初に1回だけ実行する：シートと保存フォルダを作る */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = ensureSheet_(ss, SHEET_TASKS, TASK_HEADER);
  if (tasks.getLastRow() === 1) {
    tasks.getRange(2, 1, 3, TASK_HEADER.length).setValues([
      ['sample01', '英語の母音と単語（見本）', '表示される語を、ふだんの速さで1回ずつ発音してください。', '', 1, 'heed', '/hiːd/', 10],
      ['sample01', '', '', '', 2, 'hid', '/hɪd/', 10],
      ['sample01', '', '', '', 3, 'Please call Stella.', '文はひと息で読みます', 15],
    ]);
  }
  ensureSheet_(ss, SHEET_FILES, FILES_HEADER);
  ensureSheet_(ss, SHEET_SUBS, SUBS_HEADER);
  const folder = rootFolder_();
  Logger.log('保存フォルダ: ' + folder.getUrl());
}

function ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 消されていたら作り直す */ }
  }
  const folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function safe_(s) {
  return String(s || '').trim().replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 60);
}

function fmtDate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy/MM/dd HH:mm');
}

/** 課題IDの行を集めて課題の形にする */
function getAssignment_(id) {
  if (!id) return null;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => String(r[0]).trim() === id);
  if (!rows.length) return null;
  const first = (col) => { const r = rows.find(x => String(x[col]).trim() !== ''); return r ? r[col] : ''; };
  const deadlineRaw = first(3);
  let deadline = null;
  if (deadlineRaw instanceof Date) deadline = deadlineRaw;
  else if (deadlineRaw) { const d = new Date(deadlineRaw); if (!isNaN(d)) deadline = d; }
  const items = rows
    .filter(r => String(r[5]).trim() !== '')
    .map((r, i) => ({
      no: Number(r[4]) || i + 1,
      text: String(r[5]),
      note: String(r[6] || ''),
      maxSec: Number(r[7]) || 30,
    }));
  return {
    id: id,
    title: String(first(1) || id),
    description: String(first(2) || ''),
    deadline: deadline ? fmtDate_(deadline) : '',
    closed: !!deadline && new Date() > deadline,
    items: items,
  };
}

function doGet(e) {
  try {
    const a = getAssignment_(String((e.parameter && e.parameter.a) || '').trim());
    if (!a) return json_({ ok: false, error: 'not_found' });
    return json_({ ok: true, assignment: a });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (!p.assignmentId || !p.submissionId || !p.studentId || !p.name) return json_({ ok: false, error: 'bad_request' });
    const a = getAssignment_(String(p.assignmentId));
    if (!a) return json_({ ok: false, error: 'not_found' });
    if (p.action === 'upload') return json_(upload_(p, a));
    if (p.action === 'finish') return json_(finish_(p, a));
    return json_({ ok: false, error: 'bad_request' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function upload_(p, a) {
  if (a.closed) return { ok: false, error: 'closed' };
  if (!p.data) return { ok: false, error: 'bad_request' };
  const bytes = Utilities.base64Decode(p.data);
  if (bytes.length > MAX_BYTES) return { ok: false, error: 'too_large' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filesSheet = ss.getSheetByName(SHEET_FILES);
  const sid = String(p.submissionId);
  const no = Number(p.itemNo);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // 通信の再試行で同じ項目が2回届いたら、2回目は保存しない
    const dup = filesSheet.createTextFinder(sid).matchEntireCell(true).findAll()
      .some(cell => Number(filesSheet.getRange(cell.getRow(), 6).getValue()) === no);
    if (dup) return { ok: true, duplicate: true };

    const folder = childFolder_(childFolder_(rootFolder_(), safe_(a.id)), safe_(p.studentId) + '_' + safe_(p.name));
    const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
    const fileName = [safe_(a.id), safe_(p.studentId), String(no).padStart(2, '0'), stamp].join('_') + '.wav';
    const file = folder.createFile(Utilities.newBlob(bytes, 'audio/wav', fileName));
    filesSheet.appendRow([new Date(), sid, a.id, String(p.studentId), String(p.name), no, String(p.itemText || ''),
      Number(p.durationSec) || '', Number(p.sampleRate) || '', Number(p.peak) || '', fileName, file.getUrl()]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function finish_(p, a) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sid = String(p.submissionId);
  const received = ss.getSheetByName(SHEET_FILES).createTextFinder(sid).matchEntireCell(true).findAll().length;
  const expected = Number(p.count) || a.items.length;
  const now = new Date();
  const receipt = safe_(a.id) + '-' + safe_(p.studentId) + '-' + Utilities.formatDate(now, TZ, 'MMddHHmm');
  ss.getSheetByName(SHEET_SUBS).appendRow([now, receipt, a.id, String(p.studentId), String(p.name), expected, received,
    received >= expected ? '完了' : '不足あり', sid, String(p.userAgent || '').slice(0, 200)]);
  return { ok: true, receipt: receipt, received: received, receivedAt: fmtDate_(now) };
}
