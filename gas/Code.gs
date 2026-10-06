/**
 * 音声課題の受付（Google Apps Script・スプレッドシートに紐づけて使う）
 *
 * シート「課題」     … 先生が課題を書く（1行＝1項目）。課題IDを変えれば何回分でも並べられる
 * シート「名簿」     … 受講者の学籍番号と氏名（任意。提出状況の表の行になる）
 * シート「課題一覧」  … 課題ごとの状態・提出者数・学生用リンク（自動）
 * シート「提出状況」  … 名簿×課題の提出状況の表（自動）
 * シート「受信ファイル」… 届いた録音1件ごとの記録（自動）
 * シート「提出一覧」   … 学生1人の1回の提出ごとの記録（自動）
 * 録音はドライブに次の形で保存する（課題IDと学籍番号で振り分けるので、名前の表記が変わっても分かれない）
 *   音声課題_提出/
 *     <課題ID>_<課題名>/
 *       <学籍番号>_<氏名>/            … 最新の完全な提出だけが直下に並ぶ
 *         <課題ID>_<学籍番号>_<項目番号>_<提示文>_<提出日時>.wav
 *         _以前の提出/                … 出し直す前の提出
 *         _未完了の提出/              … 通信が切れて全項目そろわなかった提出
 */

const PAGES_URL = 'https://labphonlab.github.io/voice-submit/';
const SHEET_TASKS = '課題';
const SHEET_ROSTER = '名簿';
const SHEET_OVERVIEW = '課題一覧';
const SHEET_STATUS = '提出状況';
const SHEET_FILES = '受信ファイル';
const SHEET_SUBS = '提出一覧';
const ROOT_FOLDER_NAME = '音声課題_提出';
const MAX_BYTES = 25 * 1024 * 1024; // 1項目あたりの上限
const TZ = 'Asia/Tokyo';

const TASK_HEADER = ['課題ID', '課題名', '説明', '締切', '項目番号', '提示文', '読み・補足', '最長秒数', '公開開始'];
const ROSTER_HEADER = ['学籍番号', '氏名'];
const OVERVIEW_HEADER = ['課題ID', '課題名', '項目数', '公開開始', '締切', '状態', '提出者数', '学生用リンク', 'LINE用リンク', '保存フォルダ'];
const FILES_HEADER = ['受信日時', '提出ID', '課題ID', '学籍番号', '氏名', '項目番号', '提示文', '秒数', 'サンプリング周波数', 'ピーク', 'ファイル名', 'URL', 'ファイルID'];
const SUBS_HEADER = ['完了日時', '受付番号', '課題ID', '学籍番号', '氏名', '項目数', '受信数', '状態', '提出ID', '端末'];

// ---------------------------------------------------------------- 初期設定・メニュー

/** 最初に1回だけ実行する：シート・保存フォルダ・1時間ごとの自動更新を作る */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasks = ensureSheet_(ss, SHEET_TASKS, TASK_HEADER);
  if (tasks.getLastRow() === 1) {
    tasks.getRange(2, 1, 3, TASK_HEADER.length).setValues([
      ['sample01', '英語の母音と単語（見本）', '表示される語を、ふだんの速さで1回ずつ発音してください。', '', 1, 'heed', '/hiːd/', 10, ''],
      ['sample01', '', '', '', 2, 'hid', '/hɪd/', 10, ''],
      ['sample01', '', '', '', 3, 'Please call Stella.', '文はひと息で読みます', 15, ''],
    ]);
  }
  ensureSheet_(ss, SHEET_ROSTER, ROSTER_HEADER);
  ensureSheet_(ss, SHEET_OVERVIEW, OVERVIEW_HEADER);
  ensureSheet_(ss, SHEET_STATUS, ['学籍番号', '氏名']);
  ensureSheet_(ss, SHEET_FILES, FILES_HEADER);
  ensureSheet_(ss, SHEET_SUBS, SUBS_HEADER);
  rootFolder_();
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'refresh')) {
    ScriptApp.newTrigger('refresh').timeBased().everyHours(1).create();
  }
  refresh();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('音声課題')
    .addItem('課題一覧と提出状況を更新', 'refresh')
    .addItem('新しい課題のひな形を追加', 'addAssignmentTemplate')
    .addSeparator()
    .addItem('保存フォルダを開くリンクを表示', 'showFolderLink')
    .addToUi();
}

/** 「課題」シートの末尾に、新しい課題IDで5項目分の行を足す */
function addAssignmentTemplate() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('新しい課題', '課題IDを半角英数字で入れてください（例：eng02）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const id = res.getResponseText().trim();
  if (!/^[A-Za-z0-9_-]+$/.test(id)) { ui.alert('課題IDは半角英数字と - _ だけにしてください。'); return; }
  if (allAssignments_().some(a => a.id === id)) { ui.alert('課題ID「' + id + '」はもうあります。'); return; }
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const rows = [];
  for (let i = 1; i <= 5; i++) {
    rows.push(i === 1
      ? [id, '（課題名）', '（学生への説明）', '（締切 例：2026/10/20 23:59）', 1, '（提示文）', '（読み・補足）', 10, '（公開開始。空欄ならすぐ公開）']
      : [id, '', '', '', i, '（提示文）', '', 10, '']);
  }
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, TASK_HEADER.length).setValues(rows);
  sh.activate();
  sh.setActiveRange(sh.getRange(start, 2));
  ui.alert('「課題」シートの ' + start + ' 行目から書き足しました。（ ）の中を書き換えてください。項目が足りなければ行を増やします。');
}

function showFolderLink() {
  SpreadsheetApp.getUi().alert('保存フォルダ：\n' + rootFolder_().getUrl());
}

// ---------------------------------------------------------------- 課題の読み取り

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

/** 課題のフォルダ。IDを覚えておき、課題名が変わったらフォルダ名も合わせる */
function assignmentFolder_(a) {
  const props = PropertiesService.getScriptProperties();
  const key = 'AF_' + a.id;
  const name = safe_(a.id) + '_' + safe_(a.title);
  let folder = null;
  const id = props.getProperty(key);
  if (id) {
    try { folder = DriveApp.getFolderById(id); if (folder.isTrashed()) folder = null; } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = rootFolder_().createFolder(name);
    props.setProperty(key, folder.getId());
  } else if (folder.getName() !== name) {
    folder.setName(name);
  }
  return folder;
}

function existingAssignmentFolder_(a) {
  const id = PropertiesService.getScriptProperties().getProperty('AF_' + a.id);
  if (!id) return null;
  try { const f = DriveApp.getFolderById(id); return f.isTrashed() ? null : f; } catch (e) { return null; }
}

/** 学生のフォルダ。学籍番号で探すので、氏名の書き方が変わっても同じフォルダに入る */
function studentFolder_(parent, studentId, name) {
  const prefix = safe_(studentId) + '_';
  const it = parent.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(prefix) === 0) return f;
  }
  return parent.createFolder(prefix + safe_(name));
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

function toDate_(v) {
  if (v instanceof Date) return v;
  if (v === '' || v === null || v === undefined) return null;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

/** 「課題」シートを課題IDごとにまとめる（シートに書かれた順） */
function allAssignments_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const rows = sh.getDataRange().getValues().slice(1);
  const order = [];
  const groups = {};
  rows.forEach(r => {
    const id = String(r[0]).trim();
    if (!id) return;
    if (!groups[id]) { groups[id] = []; order.push(id); }
    groups[id].push(r);
  });
  const now = new Date();
  return order.map(id => {
    const g = groups[id];
    const first = (col) => { const r = g.find(x => String(x[col]).trim() !== ''); return r ? r[col] : ''; };
    const deadline = toDate_(first(3));
    const opensAt = toDate_(first(8));
    const items = g
      .filter(r => String(r[5]).trim() !== '')
      .map((r, i) => ({ no: Number(r[4]) || i + 1, text: String(r[5]), note: String(r[6] || ''), maxSec: Number(r[7]) || 30 }))
      .sort((x, y) => x.no - y.no);
    return {
      id: id,
      title: String(first(1) || id),
      description: String(first(2) || ''),
      deadline: deadline ? fmtDate_(deadline) : '',
      opensAt: opensAt ? fmtDate_(opensAt) : '',
      notOpen: !!opensAt && now < opensAt,
      closed: !!deadline && now > deadline,
      deadlineTime: deadline ? deadline.getTime() : Infinity,
      items: items,
    };
  });
}

function getAssignment_(id) {
  if (!id) return null;
  return allAssignments_().find(a => a.id === id) || null;
}

function publicAssignment_(a) {
  return { id: a.id, title: a.title, description: a.description, deadline: a.deadline, closed: a.closed, items: a.items };
}

// ---------------------------------------------------------------- 学生画面からの呼び出し

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'list') {
      // 公開開始前の課題は出さない。受付中を締切の早い順に、そのあと締切済みを新しい順に
      const list = allAssignments_().filter(a => !a.notOpen && a.items.length);
      const open = list.filter(a => !a.closed).sort((x, y) => x.deadlineTime - y.deadlineTime);
      const closed = list.filter(a => a.closed).sort((x, y) => y.deadlineTime - x.deadlineTime);
      return json_({ ok: true, assignments: open.concat(closed).map(a => ({
        id: a.id, title: a.title, deadline: a.deadline, closed: a.closed, count: a.items.length })) });
    }
    const a = getAssignment_(String(p.a || '').trim());
    if (!a) return json_({ ok: false, error: 'not_found' });
    if (a.notOpen) return json_({ ok: false, error: 'not_open', opensAt: a.opensAt });
    return json_({ ok: true, assignment: publicAssignment_(a) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (!p.assignmentId || !p.submissionId || !p.studentId || !p.name) return json_({ ok: false, error: 'bad_request' });
    const a = getAssignment_(String(p.assignmentId));
    if (!a || a.notOpen) return json_({ ok: false, error: 'not_found' });
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

    const folder = studentFolder_(assignmentFolder_(a), p.studentId, p.name);
    // 1回の提出の全項目に同じ日時を付ける（学生が送信を押した時刻）
    const started = Number(p.startedAt) ? new Date(Number(p.startedAt)) : new Date();
    const stamp = Utilities.formatDate(started, TZ, 'yyyyMMdd-HHmm');
    const label = safe_(p.itemText).replace(/[^\p{L}\p{N}_'-]+/gu, '').slice(0, 20);
    const fileName = [safe_(a.id), safe_(p.studentId), String(no).padStart(2, '0'), label, stamp].filter(String).join('_') + '.wav';
    const file = folder.createFile(Utilities.newBlob(bytes, 'audio/wav', fileName));
    filesSheet.appendRow([new Date(), sid, a.id, String(p.studentId), String(p.name), no, String(p.itemText || ''),
      Number(p.durationSec) || '', Number(p.sampleRate) || '', Number(p.peak) || '', fileName, file.getUrl(), file.getId()]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function finish_(p, a) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sid = String(p.submissionId);
  const filesSheet = ss.getSheetByName(SHEET_FILES);
  const fileIds = filesSheet.createTextFinder(sid).matchEntireCell(true).findAll()
    .map(cell => String(filesSheet.getRange(cell.getRow(), FILES_HEADER.length).getValue()));
  const received = fileIds.length;
  const expected = Number(p.count) || a.items.length;
  const complete = received >= expected;
  const now = new Date();

  // 学生フォルダの直下には最新の完全な提出だけを残す
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const folder = studentFolder_(assignmentFolder_(a), p.studentId, p.name);
    const mine = new Set(fileIds);
    if (complete) {
      const old = [];
      const it = folder.getFiles();
      while (it.hasNext()) { const f = it.next(); if (!mine.has(f.getId())) old.push(f); }
      if (old.length) { const dest = childFolder_(folder, '_以前の提出'); old.forEach(f => f.moveTo(dest)); }
    } else if (fileIds.length) {
      const dest = childFolder_(folder, '_未完了の提出');
      fileIds.forEach(id => { try { DriveApp.getFileById(id).moveTo(dest); } catch (e) { /* 消されていたら飛ばす */ } });
    }
  } finally {
    lock.releaseLock();
  }

  const receipt = safe_(a.id) + '-' + safe_(p.studentId) + '-' + Utilities.formatDate(now, TZ, 'MMddHHmm');
  ss.getSheetByName(SHEET_SUBS).appendRow([now, receipt, a.id, String(p.studentId), String(p.name), expected, received,
    complete ? '完了' : '不足あり', sid, String(p.userAgent || '').slice(0, 200)]);
  return { ok: true, receipt: receipt, received: received, receivedAt: fmtDate_(now) };
}

// ---------------------------------------------------------------- 先生用の集計

/** 「課題一覧」と「提出状況」を作り直す（メニューと1時間ごとの自動実行から呼ぶ） */
function refresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assignments = allAssignments_();
  const subs = ss.getSheetByName(SHEET_SUBS).getDataRange().getValues().slice(1);

  // 学生×課題ごとの最良の提出（完了を優先し、同じ状態なら新しいもの）
  const best = {};
  const names = {};
  subs.forEach(r => {
    const when = r[0], aid = String(r[2]), sidNo = String(r[3]).trim(), name = String(r[4]);
    const done = r[7] === '完了';
    names[sidNo] = names[sidNo] || name;
    const key = sidNo + '\t' + aid;
    const cur = best[key];
    if (!cur || (done && !cur.done) || (done === cur.done && when > cur.when)) {
      best[key] = { when: when, done: done, received: r[6], expected: r[5] };
    }
  });

  // 課題一覧
  const overview = assignments.map(a => {
    const submitters = Object.keys(best).filter(k => k.split('\t')[1] === a.id && best[k].done).length;
    const state = a.notOpen ? '公開前' : (a.closed ? '締切済み' : '受付中');
    const link = PAGES_URL + '?a=' + encodeURIComponent(a.id);
    const folder = existingAssignmentFolder_(a);
    return [a.id, a.title, a.items.length, a.opensAt, a.deadline, state, submitters,
      link, link + '&openExternalBrowser=1', folder ? folder.getUrl() : '（まだ提出なし）'];
  });
  const ov = ss.getSheetByName(SHEET_OVERVIEW);
  ov.getRange(2, 1, Math.max(ov.getMaxRows() - 1, 1), OVERVIEW_HEADER.length).clearContent();
  if (overview.length) ov.getRange(2, 1, overview.length, OVERVIEW_HEADER.length).setValues(overview);

  // 提出状況：名簿の学生を上に、名簿にない提出者を下に「名簿外」として並べる
  const roster = ss.getSheetByName(SHEET_ROSTER).getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim() !== '').map(r => [String(r[0]).trim(), String(r[1])]);
  const inRoster = new Set(roster.map(r => r[0]));
  const extra = Object.keys(names).filter(id => !inRoster.has(id)).sort().map(id => [id, names[id] + '（名簿外）']);
  const students = roster.concat(extra);
  const shown = assignments.filter(a => !a.notOpen);
  const header = ['学籍番号', '氏名'].concat(shown.map(a => a.id + '\n' + a.title)).concat(['提出数']);
  const rows = students.map(s => {
    let n = 0;
    const cells = shown.map(a => {
      const b = best[s[0] + '\t' + a.id];
      if (!b) return a.closed ? '未提出' : '';
      if (b.done) { n++; return '完了 ' + Utilities.formatDate(b.when, TZ, 'MM/dd HH:mm'); }
      return '不足 ' + b.received + '/' + b.expected;
    });
    return s.concat(cells).concat([n + ' / ' + shown.length]);
  });
  const st = ss.getSheetByName(SHEET_STATUS);
  st.clear();
  st.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setWrap(true);
  st.setFrozenRows(1);
  st.setFrozenColumns(2);
  if (rows.length) st.getRange(2, 1, rows.length, header.length).setValues(rows);
  st.getRange(1, 1, 1, 1).setNote('最終更新 ' + fmtDate_(new Date()));
}
