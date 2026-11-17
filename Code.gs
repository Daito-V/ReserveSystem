/** 会議室予約システム - Google Apps Script backend */
const SHEETS = {
  rooms: ['id', 'name', 'capacity', 'grace_minutes', 'is_active'],
  users: ['id', 'name', 'status', 'role'],
  reservations: ['id', 'room_id', 'start_at', 'end_at', 'user_id', 'status', 'created_at'],
  cancellations: ['id', 'reservation_id', 'canceled_by', 'canceled_at', 'is_proxy', 'reason']
};
const LABELS = { BOOKED: '予約済', IN_USE: '利用中', USED: '利用済', CANCELED: 'キャンセル済', AUTO_CANCELED: '自動キャンセル済' };
const FIXED_ROOMS = [
  [1, '第1会議室', 0, 15, true], [2, '第2会議室', 0, 15, true],
  [3, '第3会議室', 0, 15, true], [4, '第4会議室', 0, 15, true],
  [5, '役員会議室', 0, 30, true]
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('会議室予約システム');
}

/** 初回実行用。シートとヘッダーを作り、roomsが空なら固定5室を登録します。usersは既存社員データを入力してください。 */
function initializeDatabase() {
  const ss = getSpreadsheet_();
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
    else sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
    sh.setFrozenRows(1);
  });
  const rooms = ss.getSheetByName('rooms');
  if (rooms.getLastRow() === 1) rooms.getRange(2, 1, FIXED_ROOMS.length, FIXED_ROOMS[0].length).setValues(FIXED_ROOMS);
  return 'シートを初期化しました。usersシートに社員データを登録してください。';
}

function getInitialData() {
  const rooms = rows_('rooms').filter(r => truth_(r.is_active)).map(r => ({ id: num_(r.id), name: String(r.name) }));
  const users = rows_('users').filter(r => String(r.status) === 'ACTIVE').map(r => ({ id: num_(r.id), name: String(r.name), role: String(r.role) }));
  return { rooms, users };
}

function getReservations(dateString) {
  if (!dateString) throw new Error('日付を指定してください。');
  const day = String(dateString).slice(0, 10);
  const rows = rows_('reservations').filter(r => {
    const d = dateValue_(r.start_at);
    return d && Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') === day;
  });
  const rooms = indexById_('rooms'), users = indexById_('users');
  return rows.map(r => ({
    id: num_(r.id), room_id: num_(r.room_id), room_name: rooms[String(num_(r.room_id))]?.name || '不明な会議室',
    start_at: iso_(r.start_at), end_at: iso_(r.end_at), user_id: num_(r.user_id),
    user_name: users[String(num_(r.user_id))]?.name || '不明な利用者', status: String(r.status), status_label: LABELS[String(r.status)] || String(r.status)
  })).sort((a,b) => a.room_name.localeCompare(b.room_name, 'ja') || new Date(a.start_at)-new Date(b.start_at));
}

function createReservation(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const roomId = num_(input && input.room_id), userId = num_(input && input.user_id);
    const start = dateValue_(input && input.start_at), end = dateValue_(input && input.end_at);
    // B-1
    if (!roomId || !start || !end || !userId) throw new Error('会議室、開始日時、終了日時、予約者は必須です。');
    const room = findById_('rooms', roomId);
    const user = findById_('users', userId);
    if (!room) throw new Error('指定された会議室が見つかりません。');
    // B-2
    if (start.getTime() >= end.getTime()) throw new Error('開始日時は終了日時より前にしてください。');
    // B-3
    if (!user || String(user.status) !== 'ACTIVE') throw new Error('予約者はACTIVEの社員を指定してください。');
    // B-4/B-5
    const reservations = rows_('reservations');
    if (reservations.some(r => num_(r.room_id) === roomId && !['CANCELED','AUTO_CANCELED'].includes(String(r.status)) && start < dateValue_(r.end_at) && end > dateValue_(r.start_at))) {
      throw new Error('同一会議室の予約時間が重複しています。');
    }
    // B-6: 日付は予約開始日時の暦日、スクリプトのタイムゾーン基準
    const dateKey = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const count = reservations.filter(r => num_(r.user_id) === userId && dateValue_(r.start_at) && Utilities.formatDate(dateValue_(r.start_at), Session.getScriptTimeZone(), 'yyyy-MM-dd') === dateKey && !['CANCELED','AUTO_CANCELED'].includes(String(r.status))).length;
    if (count >= 3) throw new Error('予約は1人あたり1日3件までです。');
    const id = nextId_('reservations');
    append_('reservations', [id, roomId, start, end, userId, 'BOOKED', new Date()]);
    return { id: id, message: '予約しました。' };
  } finally { lock.releaseLock(); }
}

function cancelOwnReservation(reservationId, userId, reason) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const reservation = findById_('reservations', num_(reservationId));
    if (!reservation) throw new Error('対象の予約が見つかりません。');
    const user = findById_('users', num_(userId));
    if (!user || num_(reservation.user_id) !== num_(userId)) throw new Error('本人の予約ではありません。');
    if (String(reservation.status) !== 'BOOKED') throw new Error('この予約はキャンセルできる状態ではありません。');
    const boundary = dateValue_(reservation.start_at).getTime() - 24 * 60 * 60 * 1000;
    if (Date.now() > boundary) throw new Error('予約開始時刻の24時間前を過ぎているため、本人はキャンセルできません。');
    finishCancellation_(reservation, userId, false, reason || '');
    return 'キャンセルしました。';
  } finally { lock.releaseLock(); }
}

function cancelAsAdmin(reservationId, actorId, reason) {
  // F-4: 判定順序・文言を要件通りに実施
  const rid = num_(reservationId), aid = num_(actorId);
  const reservation = rid ? findById_('reservations', rid) : null;
  if (!reservation) throw new Error('対象の予約が見つかりません。');
  const actor = aid ? findById_('users', aid) : null;
  if (!actor || String(actor.role) !== 'ADMIN') throw new Error('代理キャンセルの権限がありません。');
  const cleaned = String(reason == null ? '' : reason).trim();
  if (!cleaned) throw new Error('代理キャンセルの理由を入力してください。');
  if (cleaned.length > 200) throw new Error('理由は200文字以内で入力してください。');
  if (String(reservation.status) !== 'BOOKED') throw new Error('この予約はキャンセルできる状態ではありません。');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const current = findById_('reservations', rid);
    if (!current || String(current.status) !== 'BOOKED') throw new Error('この予約はキャンセルできる状態ではありません。');
    finishCancellation_(current, aid, true, cleaned);
    return 'キャンセルしました。';
  } finally { lock.releaseLock(); }
}

function getCancellationHistory(reservationId) {
  const rid = num_(reservationId);
  const reservation = rid ? findById_('reservations', rid) : null;
  if (!reservation) throw new Error('対象の予約が見つかりません。');
  if (!['CANCELED','AUTO_CANCELED'].includes(String(reservation.status))) return [];
  const users = indexById_('users');
  return rows_('cancellations').filter(r => num_(r.reservation_id) === rid).map(r => ({
    canceled_by: num_(r.canceled_by), canceled_by_name: users[String(num_(r.canceled_by))]?.name || '不明な利用者',
    canceled_at: iso_(r.canceled_at), reason: String(r.reason || '').trim() || '理由の記載なし', is_proxy: truth_(r.is_proxy)
  })).sort((a,b) => new Date(a.canceled_at)-new Date(b.canceled_at));
}

function finishCancellation_(reservation, actorId, proxy, reason) {
  const ss = getSpreadsheet_(), sh = ss.getSheetByName('reservations');
  const row = findRowNumber_('reservations', num_(reservation.id));
  if (!row) throw new Error('対象の予約が見つかりません。');
  sh.getRange(row, 6).setValue('CANCELED');
  append_('cancellations', [nextId_('cancellations'), num_(reservation.id), num_(actorId), new Date(), proxy, reason]);
}
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートに紐付けるか、SPREADSHEET_IDを設定してください。');
  return ss;
}
function rows_(name) { const sh=getSpreadsheet_().getSheetByName(name); if(!sh || sh.getLastRow()<2) return []; const v=sh.getDataRange().getValues(); return v.slice(1).filter(r=>r.some(x=>x!==''&&x!==null)).map(row=>Object.fromEntries(SHEETS[name].map((k,i)=>[k,row[i]]))); }
function findById_(name,id) { return rows_(name).find(r=>num_(r.id)===num_(id))||null; }
function findRowNumber_(name,id) { const all=rows_(name); const i=all.findIndex(r=>num_(r.id)===num_(id)); return i<0?null:i+2; }
function indexById_(name) { return Object.fromEntries(rows_(name).map(r=>[String(num_(r.id)),r])); }
function append_(name,row) { getSpreadsheet_().getSheetByName(name).appendRow(row); }
function nextId_(name) { return rows_(name).reduce((m,r)=>Math.max(m,num_(r.id)||0),0)+1; }
function num_(v) { if(v===null||v===undefined||String(v).trim()==='') return NaN; const n=Number(v); return Number.isFinite(n)?n:NaN; }
function truth_(v) { return v===true || String(v).toUpperCase()==='TRUE'; }
function dateValue_(v) { if(v instanceof Date && !isNaN(v)) return v; if(v===null||v===undefined||v==='') return null; const s=String(v); if(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) { try { return Utilities.parseDate(s, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"); } catch(e) { return null; } } const d=new Date(v); return isNaN(d)?null:d; }
function iso_(v) { const d=dateValue_(v); return d?d.toISOString():''; }
