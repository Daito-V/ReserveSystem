const SHEETS = {
  rooms: 'rooms',
  users: 'users',
  reservations: 'reservations',
  cancellations: 'cancellations',
};

const HEADERS = {
  rooms: ['id', 'name', 'capacity', 'grace_minutes', 'is_active'],
  users: ['id', 'name', 'status', 'role'],
  reservations: ['id', 'room_id', 'start_at', 'end_at', 'user_id', 'status', 'created_at'],
  cancellations: ['id', 'reservation_id', 'canceled_by', 'canceled_at', 'is_proxy', 'reason'],
};

const ROOM_SEED = [
  [1, '第1会議室', 8, 15, true],
  [2, '第2会議室', 8, 15, true],
  [3, '第3会議室', 10, 15, true],
  [4, '第4会議室', 12, 15, true],
  [5, '役員会議室', 16, 30, true],
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('会議室予約システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 初回のみ実行。既存データを消さず、シートとヘッダー、会議室5室を準備する。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (key) {
    const name = SHEETS[key];
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS[key].length).setValues([HEADERS[key]]);
      sheet.setFrozenRows(1);
    }
  });
  const rooms = ss.getSheetByName(SHEETS.rooms);
  if (rooms.getLastRow() === 1) rooms.getRange(2, 1, ROOM_SEED.length, ROOM_SEED[0].length).setValues(ROOM_SEED);
  return 'セットアップが完了しました。users シートに社員を登録してください。';
}

function getInitialData() {
  return {
    rooms: readRows_('rooms').filter(function (r) { return truthy_(r.is_active); }),
    users: readRows_('users').filter(function (u) { return u.status === 'ACTIVE'; }),
  };
}

function listReservations(dateText) {
  const day = String(dateText || '');
  // Sheets I/O is the expensive part. Read each table once, then resolve names in memory.
  const roomsById = new Map(readRows_('rooms').map(function (room) { return [Number(room.id), room]; }));
  const usersById = new Map(readRows_('users').map(function (user) { return [Number(user.id), user]; }));
  return readRows_('reservations')
    .filter(function (r) { return localDate_(r.start_at) === day; })
    .map(function (r) {
      const room = roomsById.get(Number(r.room_id));
      const user = usersById.get(Number(r.user_id));
      return Object.assign({}, r, {
        room_name: room ? room.name : '不明な会議室',
        user_name: user ? user.name : '不明な予約者',
        start_at: toIso_(r.start_at),
        end_at: toIso_(r.end_at),
        status_label: statusLabel_(r.status),
      });
    })
    .sort(function (a, b) {
      return Number(a.room_id) - Number(b.room_id) || new Date(a.start_at) - new Date(b.start_at);
    });
}

function createReservation(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const roomId = Number(input && input.room_id);
    const userId = Number(input && input.user_id);
    const start = new Date(input && input.start_at);
    const end = new Date(input && input.end_at);
    if (!roomId || !input.start_at || !input.end_at || !userId) return fail_('会議室・開始日時・終了日時・予約者を指定してください。');
    if (isNaN(start.getTime()) || isNaN(end.getTime()) || start >= end) return fail_('開始日時は終了日時より前にしてください。');
    const user = findById_('users', userId);
    if (!user || user.status !== 'ACTIVE') return fail_('予約者は在籍中の社員を指定してください。');

    const reservations = readRows_('reservations');
    const overlaps = reservations.some(function (r) {
      if (Number(r.room_id) !== roomId || ['CANCELED', 'AUTO_CANCELED'].indexOf(r.status) >= 0) return false;
      const oldStart = new Date(r.start_at), oldEnd = new Date(r.end_at);
      return start < oldEnd && end > oldStart;
    });
    if (overlaps) return fail_('同じ会議室に重複する予約があります。');

    const day = localDate_(start);
    const count = reservations.filter(function (r) {
      return Number(r.user_id) === userId && localDate_(r.start_at) === day;
    }).length;
    if (count >= 3) return fail_('予約は1人あたり1日3件までです。');

    const id = nextId_('reservations');
    appendRow_('reservations', [id, roomId, start, end, userId, 'BOOKED', new Date()]);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function cancelOwn(reservationId, actorId, reason) {
  const reservation = findById_('reservations', Number(reservationId));
  if (!reservation) return fail_('対象の予約が見つかりません。');
  if (Number(reservation.user_id) !== Number(actorId)) return fail_('自分が登録した予約のみキャンセルできます。');
  if (reservation.status !== 'BOOKED') return fail_('この予約はキャンセルできる状態ではありません。');
  if (new Date(reservation.start_at).getTime() - Date.now() < 24 * 60 * 60 * 1000) return fail_('予約開始時刻の24時間前を過ぎた予約はキャンセルできません。');
  const trimmed = String(reason || '').trim();
  if (trimmed.length > 200) return fail_('理由は200文字以内で入力してください。');
  const result = recordCancellation_(reservation, Number(actorId), false, trimmed);
  return result.ok ? { ok: true, message: 'キャンセルしました。' } : result;
}

function proxyCancel(reservationId, actorId, reason) {
  const reservation = findById_('reservations', Number(reservationId));
  if (!reservation) return fail_('対象の予約が見つかりません。');
  const actor = findById_('users', Number(actorId));
  if (!actor || actor.role !== 'ADMIN') return fail_('代理キャンセルの権限がありません。');
  const trimmed = String(reason || '').trim();
  if (!trimmed) return fail_('代理キャンセルの理由を入力してください。');
  if (trimmed.length > 200) return fail_('理由は200文字以内で入力してください。');
  if (reservation.status !== 'BOOKED') return fail_('この予約はキャンセルできる状態ではありません。');
  const result = recordCancellation_(reservation, Number(actorId), true, trimmed);
  return result.ok ? { ok: true, message: 'キャンセルしました。' } : result;
}

function getReservationDetail(reservationId) {
  const reservation = findById_('reservations', Number(reservationId));
  if (!reservation) return null;
  const room = findById_('rooms', reservation.room_id);
  const user = findById_('users', reservation.user_id);
  const cancellation = readRows_('cancellations').find(function (c) { return Number(c.reservation_id) === Number(reservation.id); });
  const canceler = cancellation && findById_('users', cancellation.canceled_by);
  return {
    id: reservation.id,
    room_name: room ? room.name : '不明な会議室',
    user_name: user ? user.name : '不明な予約者',
    start_at: toIso_(reservation.start_at),
    end_at: toIso_(reservation.end_at),
    status: reservation.status,
    status_label: statusLabel_(reservation.status),
    cancellation: cancellation ? {
      canceled_by_name: canceler ? canceler.name : '不明な実行者',
      canceled_at: toIso_(cancellation.canceled_at),
      is_proxy: truthy_(cancellation.is_proxy),
      reason: cancellation.reason || '（理由の記載なし）',
    } : null,
  };
}

function recordCancellation_(reservation, actorId, isProxy, reason) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const current = findById_('reservations', Number(reservation.id));
    if (!current || current.status !== 'BOOKED') return fail_('この予約はキャンセルできる状態ではありません。');
    updateReservationStatus_(Number(current.id), 'CANCELED');
    appendRow_('cancellations', [nextId_('cancellations'), Number(current.id), actorId, new Date(), isProxy, reason]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function updateReservationStatus_(id, status) {
  const sheet = sheet_('reservations');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === id) { sheet.getRange(i + 1, 6).setValue(status); return; }
  }
}

function readRows_(key) {
  const values = sheet_(key).getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(function (row) { return row.some(function (v) { return v !== ''; }); })
    .map(function (row) {
      const out = {};
      headers.forEach(function (h, i) { out[h] = row[i]; });
      return out;
    });
}

function findById_(key, id) {
  return readRows_(key).find(function (r) { return Number(r.id) === Number(id); }) || null;
}
function nextId_(key) { return readRows_(key).reduce(function (max, row) { return Math.max(max, Number(row.id) || 0); }, 0) + 1; }
function appendRow_(key, row) { sheet_(key).appendRow(row); }
function sheet_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[key]);
  if (!sheet) throw new Error('先に setup() を実行してください。');
  return sheet;
}
function fail_(message) { return { ok: false, message: message }; }
function truthy_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function toIso_(value) { return value instanceof Date ? value.toISOString() : new Date(value).toISOString(); }
function localDate_(value) {
  const d = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function statusLabel_(status) {
  return ({ BOOKED: '予約済', IN_USE: '利用中', USED: '利用済', CANCELED: 'キャンセル済', AUTO_CANCELED: '自動キャンセル済' })[status] || status;
}
