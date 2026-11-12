const fs = require('node:fs');
const vm = require('node:vm');

class Sheet {
  constructor(rows) { this.rows = rows; this.readCount = 0; }
  getDataRange() { return { getValues: () => { this.readCount++; return this.rows.map(row => row.slice()); } }; }
  appendRow(row) { this.rows.push(row.slice()); }
  getRange(row, col) { return { setValue: value => { this.rows[row - 1][col - 1] = value; } }; }
}

const headers = {
  users: ['id', 'name', 'status', 'role'],
  rooms: ['id', 'name', 'capacity', 'grace_minutes', 'is_active'],
  reservations: ['id', 'room_id', 'start_at', 'end_at', 'user_id', 'status', 'created_at'],
  cancellations: ['id', 'reservation_id', 'canceled_by', 'canceled_at', 'is_proxy', 'reason'],
};
const data = {
  users: [[1, '社員1', 'ACTIVE', 'MEMBER'], [2, '社員2', 'ACTIVE', 'MEMBER'], [3, '総務管理者', 'ACTIVE', 'ADMIN']],
  rooms: [[1, '第1会議室', 8, 15, true]],
  reservations: [
    [1, 1, new Date('2035-05-01T01:00:00Z'), new Date('2035-05-01T02:00:00Z'), 1, 'BOOKED', new Date()],
    [2, 1, new Date('2035-05-01T03:00:00Z'), new Date('2035-05-01T04:00:00Z'), 2, 'BOOKED', new Date()],
    [3, 1, new Date('2035-05-01T05:00:00Z'), new Date('2035-05-01T06:00:00Z'), 2, 'CANCELED', new Date()],
    [4, 1, new Date('2035-05-01T07:00:00Z'), new Date('2035-05-01T08:00:00Z'), 2, 'IN_USE', new Date()],
  ],
  cancellations: [],
};
const sheets = {};
for (const [key, cols] of Object.entries(headers)) sheets[key] = new Sheet([cols, ...data[key]]);
const names = { rooms: 'rooms', users: 'users', reservations: 'reservations', cancellations: 'cancellations' };
const context = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => sheets[Object.keys(names).find(key => names[key] === name)] }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: { formatDate: (d, _tz, format) => format === 'yyyy-MM-dd' ? d.toISOString().slice(0, 10) : d.toISOString() },
  Session: { getScriptTimeZone: () => 'UTC' },
  Date, Math, Number, String, Object, Array, isNaN,
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('Code.gs', 'utf8'), context, { filename: 'Code.gs' });

const expectedFailure = (result, message) => result && result.ok === false && result.message === message;
const results = [];
const run = (name, test) => {
  try { const detail = test(); results.push({ case: name, result: detail ? 'PASS' : 'FAIL' }); }
  catch (error) { results.push({ case: name, result: 'FAIL', detail: error.message }); }
};

run('1 Admin proxy cancellation succeeds', () => {
  const r = context.proxyCancel('1', '3', '来客のため');
  return r.ok === true && r.message === 'キャンセルしました。';
});
run('2 Missing reservation', () => expectedFailure(context.proxyCancel('99', '3', '理由あり'), '対象の予約が見つかりません。'));
run('3 Non-admin rejected', () => expectedFailure(context.proxyCancel('2', '1', '理由あり'), '代理キャンセルの権限がありません。'));
run('4 Empty reason rejected', () => expectedFailure(context.proxyCancel('2', '3', ''), '代理キャンセルの理由を入力してください。'));
run('5 Whitespace-only reason rejected', () => expectedFailure(context.proxyCancel('2', '3', '  　 '), '代理キャンセルの理由を入力してください。'));
run('6 201 characters rejected', () => expectedFailure(context.proxyCancel('2', '3', 'あ'.repeat(201)), '理由は200文字以内で入力してください。'));
run('7 Exactly 200 characters accepted', () => {
  const r = context.proxyCancel('2', '3', 'あ'.repeat(200));
  return r.ok === true && r.message === 'キャンセルしました。';
});
run('8 CANCELED reservation rejected', () => expectedFailure(context.proxyCancel('3', '3', '理由あり'), 'この予約はキャンセルできる状態ではありません。'));
run('9 IN_USE reservation rejected', () => expectedFailure(context.proxyCancel('4', '3', '理由あり'), 'この予約はキャンセルできる状態ではありません。'));
run('10 Cancellation row and stored values', () => {
  const rows = sheets.cancellations.rows.slice(1);
  const r1 = rows.find(row => Number(row[1]) === 1);
  const r2 = rows.find(row => Number(row[1]) === 2);
  const res1 = sheets.reservations.rows.find(row => Number(row[0]) === 1);
  const res2 = sheets.reservations.rows.find(row => Number(row[0]) === 2);
  return rows.length === 2 && res1[5] === 'CANCELED' && res2[5] === 'CANCELED' &&
    r1[4] === true && r1[5] === '来客のため' && r2[4] === true && r2[5] === 'あ'.repeat(200);
});
run('11 List 100 reservations with one read per sheet', () => {
  for (let id = 10; id < 110; id++) {
    sheets.reservations.appendRow([id, 1, new Date('2035-05-02T01:00:00Z'), new Date('2035-05-02T02:00:00Z'), 1, 'BOOKED', new Date()]);
  }
  const before = Object.values(sheets).reduce((sum, sheet) => sum + sheet.readCount, 0);
  const result = context.listReservations('2035-05-02');
  const reads = Object.values(sheets).reduce((sum, sheet) => sum + sheet.readCount, 0) - before;
  return result.length === 100 && reads === 3;
});

console.log(JSON.stringify({ passed: results.filter(r => r.result === 'PASS').length, total: results.length, results }, null, 2));
if (results.some(r => r.result !== 'PASS')) process.exitCode = 1;
