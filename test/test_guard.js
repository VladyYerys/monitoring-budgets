// Запобіжник: шаблон змінює лише порожню таблицю або вже позначену як шаблон; таблицю з «m» і «y» не чіпає.
const assert = require('assert');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('empty spreadsheet → allowed and marked as template; second call: still allowed', () => {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  assert.strictEqual(env.ctx.guard_(env.ss, true), true);
  assert.strictEqual(env.docProps.getProperty('mbs'), 'monitoring-budgets-template');
  env.ss.getSheetByName('Аркуш1').getRange('A1').setValue('щось');
  assert.strictEqual(env.ctx.guard_(env.ss), true);
});

t('a table with sheets «m» and «y» → refused, nothing written, even if marked', () => {
  const env = makeEnv();
  ['Журнал', 'm', 'y'].forEach(s => env.ss.insertSheet(s));
  env.docProps.setProperty('mbs', 'monitoring-budgets-template');
  assert.throws(() => env.ctx.guard_(env.ss, true), /особисту таблицю/);
  const before = env.ss.getSheets().map(s => s.getName()).join('|');
  env.ctx.setupTemplate();
  assert.strictEqual(env.ss.getSheets().map(s => s.getName()).join('|'), before);
  assert.ok(/особисту таблицю/.test(env.calls.dialogs.pop().html));
  assert.strictEqual(env.ss.getSheetByName('Журнал').getLastRow(), 0);
});

t('any other non-empty, unmarked table → refused', () => {
  const env = makeEnv();
  env.ss.insertSheet('Мій облік').getRange('A1').setValue('дані');
  assert.throws(() => env.ctx.guard_(env.ss, true), /не шаблон/);
  assert.strictEqual(env.docProps.getProperty('mbs'), null);
  assert.throws(() => env.ctx.historyWrite_(env.ss, '2026-10'), /не шаблон/);
});

// Усі публічні функції (пункти меню, розклади, onOpen / onEdit): у чужій таблиці нічого не змінюють.
const { SRC } = require('./source');
const PUBLIC = Array.from(SRC.matchAll(/^function ([A-Za-z][A-Za-z0-9]*[A-Za-z0-9])\(/gm)).map(m => m[1]).filter(f => !/_$/.test(f));
const snap = env => JSON.stringify({
  sheets: env.ss.getSheets().map(s => [s.getName(), Array.from(s.cells.entries()), s.prots.length, s.cf.length, (s.charts || []).length]),
  doc: env.docProps.getProperties(), user: env.userProps.getProperties(), trg: env.triggers.length, files: env.files.length, cals: env.cals.length,
  fetch: env.calls.fetch.length, toasts: env.calls.toasts.length });

function hostile(kind) {
  const env = makeEnv({ promptText: '2026-10', alertAnswer: 'YES', mono: () => ({ code: 200, body: { ok: true, result: [] } }) });
  if (kind === 'personal') {
    ['Журнал', 'm', 'y', 'Налаштування'].forEach(s => env.ss.insertSheet(s));
    env.ss.getSheetByName('Журнал').getRange('A5:C5').setValues([['2026-11-01', 100, 'Їжа']]);
    env.docProps.setProperty('mbs', 'monitoring-budgets-template');                 // навіть з міткою
    env.ss.getSheetByName('Налаштування').getRange('A40').setValue('monitoring-budgets-template');
  } else {
    env.ss.insertSheet('Мій облік').getRange('A1:B1').setValues([['дані', 1]]);
  }
  env.userProps.setProperties({ mono_token: 'x'.repeat(40), mono_account: 'acc', tg_token: '1:' + 'y'.repeat(35), tg_chat: '1' });
  return env;
}

['personal', 'other'].forEach(kind => t('in a ' + kind + ' table every public function (' + PUBLIC.length + ') changes nothing: no cells, triggers, requests, copies, calendars', () => {
  const env = hostile(kind), before = snap(env);
  assert.ok(PUBLIC.length >= 30, PUBLIC.join(','));
  const ev = { range: env.ss.getSheetByName(kind === 'personal' ? 'Журнал' : 'Мій облік').getRange('B5') };
  PUBLIC.forEach(fn => {
    assert.strictEqual(typeof env.ctx[fn], 'function', fn);
    try { env.ctx[fn](ev); } catch (e) { /* відмова з помилкою теж ок: головне, що нічого не змінено */ }
    assert.strictEqual(snap(env), before, fn + ' змінив щось у таблиці «' + kind + '»');
  });
  if (kind === 'personal') assert.ok(!env.calls.menu, 'у особистій таблиці меню шаблону не з\'являється');
}));

t('trigger handlers have their own names (mbs…): the template never answers triggers of other code', () => {
  const env = makeEnv(), trg = env.ctx.TRG || require('vm').runInContext('TRG', env.ctx);
  Object.keys(trg).forEach(k => assert.ok(/^mbs[A-Z]/.test(trg[k]), trg[k]));
  ['sortIfIdle', 'dailyImport', 'nightlyBackup', 'morningReport', 'importRetry', 'morningRetry', 'editWatch', 'sortSoon']
    .forEach(fn => assert.strictEqual(typeof env.ctx[fn], 'undefined', fn));
});

console.log(n + ' tests passed');
