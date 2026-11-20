/** 代理キャンセル専用のGoogle Apps Script */
const SHEETS = {
  rooms: ['id', 'name', 'capacity', 'grace_minutes', 'is_active'],
  users: ['id', 'name', 'status', 'role'],
  reservations: ['id', 'room_id', 'start_at', 'end_at', 'user_id', 'status', 'created_at'],
  cancellations: ['id', 'reservation_id', 'canceled_by', 'canceled_at', 'is_proxy', 'reason']
};

const FIXED_ROOMS = [
  [1, '第1会議室', 0, 15, true],
  [2, '第2会議室', 0, 15, true],
  [3, '第3会議室', 0, 15, true],
  [4, '第4会議室', 0, 15, true],
  [5, '役員会議室', 0, 30, true]
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('会議室の代理キャンセル');
}

/** 初回実行用。必要なテーブルシートとヘッダーを用意します。 */
function initializeDatabase() {
  const spreadsheet = getSpreadsheet_();

  Object.keys(SHEETS).forEach(name => {
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
    }

    sheet.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
    sheet.setFrozenRows(1);
  });

  const roomsSheet = spreadsheet.getSheetByName('rooms');
  if (roomsSheet.getLastRow() === 1) {
    roomsSheet
      .getRange(2, 1, FIXED_ROOMS.length, FIXED_ROOMS[0].length)
      .setValues(FIXED_ROOMS);
  }

  return 'シートを初期化しました。usersシートに社員データを登録してください。';
}

/**
 * 総務部管理者による代理キャンセル。
 * 判定順序とエラーメッセージは要件の指定に従います。
 */
function cancelAsAdmin(reservationId, actorId, reason) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const reservationKey = toNumber_(reservationId);
    const reservationMatch = Number.isFinite(reservationKey)
      ? findByIdWithRow_('reservations', reservationKey)
      : null;
    const reservation = reservationMatch ? reservationMatch.record : null;

    // 1. 予約ID未指定、または対象予約なし
    if (!reservationMatch) {
      throw new Error('対象の予約が見つかりません。');
    }

    const actorKey = toNumber_(actorId);
    const actor = Number.isFinite(actorKey)
      ? findById_('users', actorKey)
      : null;

    // 2. 実行者の権限
    if (!actor || String(actor.role) !== 'ADMIN') {
      throw new Error('代理キャンセルの権限がありません。');
    }

    const cleanedReason = String(reason == null ? '' : reason).trim();

    // 3. 理由の必須確認
    if (!cleanedReason) {
      throw new Error('代理キャンセルの理由を入力してください。');
    }

    // 4. 理由の文字数
    if (cleanedReason.length > 200) {
      throw new Error('理由は200文字以内で入力してください。');
    }

    // 5. 予約状態
    if (String(reservation.status) !== 'BOOKED') {
      throw new Error('この予約はキャンセルできる状態ではありません。');
    }

    recordProxyCancellation_(
      reservation,
      reservationMatch.rowNumber,
      actorKey,
      cleanedReason
    );
    return 'キャンセルしました。';
  } finally {
    lock.releaseLock();
  }
}

function recordProxyCancellation_(reservation, reservationRow, actorId, reason) {
  const spreadsheet = getSpreadsheet_();
  spreadsheet.getSheetByName('reservations')
    .getRange(reservationRow, 6)
    .setValue('CANCELED');

  const cancellationId = nextId_('cancellations');
  spreadsheet.getSheetByName('cancellations').appendRow([
    cancellationId,
    toNumber_(reservation.id),
    actorId,
    new Date(),
    true,
    reason
  ]);
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty('SPREADSHEET_ID');
  const spreadsheet = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('スプレッドシートに紐付けるか、SPREADSHEET_IDを設定してください。');
  }

  return spreadsheet;
}

function rows_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const dataRows = values.slice(1);
  const nonEmptyRows = dataRows.filter(row =>
    row.some(value => value !== '' && value !== null)
  );

  return nonEmptyRows.map(row => {
    const record = {};
    SHEETS[sheetName].forEach((columnName, index) => {
      record[columnName] = row[index];
    });
    return record;
  });
}

function findById_(sheetName, id) {
  const records = rows_(sheetName);
  return records.find(record => toNumber_(record.id) === id) || null;
}

function findByIdWithRow_(sheetName, id) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, SHEETS[sheetName].length).getValues();

  for (let index = 0; index < values.length; index += 1) {
    const row = values[index];
    const record = {};
    SHEETS[sheetName].forEach((columnName, columnIndex) => {
      record[columnName] = row[columnIndex];
    });

    if (toNumber_(record.id) === id) {
      return {
        record: record,
        rowNumber: index + 2
      };
    }
  }

  return null;
}

function nextId_(sheetName) {
  const records = rows_(sheetName);
  const largestId = records.reduce((largest, record) => {
    return Math.max(largest, toNumber_(record.id) || 0);
  }, 0);
  return largestId + 1;
}

function toNumber_(value) {
  if (value === null || value === undefined || String(value).trim() === '') {
    return NaN;
  }

  const number = Number(value);
  return Number.isFinite(number) ? number : NaN;
}
