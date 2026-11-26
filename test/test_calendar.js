// Календар платежів: плани журналу → події; факт → «✅»; рядок прибрали → подію прибрано; чужі календарі й події не чіпаються.
const assert = require('assert');
const { makeEnv } = require('./gasmock');
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

function setup() {
  const env = makeEnv();
  env.ss.insertSheet('Аркуш1');
  env.ctx.setupTemplate();
  return env;
}
const D = k => new Date(Date.UTC(+k.slice(0, 4), +k.slice(5, 7) - 1, +k.slice(8, 10), 10));

t('plans (status «план», future date without status) become all-day events with two reminders; income and past are skipped', () => {
  const env = setup(), c = env.ctx, today = env.Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  // чужий календар з дуже схожою назвою (як в особистій таблиці): не чіпати
  const foreign = env.ctx.CalendarApp.createCalendar('Monitoring Budget: платежі', {});
  const fe = foreign.createAllDayEvent('🔁 Оренда · 9 000 ₴', D(c.addDays_(today, 3)), {}); fe.setTag('mb', 'R|Оренда|x');
  const j = env.ss.getSheetByName('Журнал');
  const put = (r, vals, type, ord) => { j.getRange(r, 1, 1, 8).setValues([vals]); j.put(r, 9, type, j.cell(r, 9).f); j.getRange(r, 31).setValue(ord); };
  put(5, [D(c.addDays_(today, 5)), 700, 'Комуналка', 'рахунок', 'план', '', '', ''], 'Витрати', 111);
  put(6, [D(c.addDays_(today, 9)), 300, 'Підписки', '', '', '', '', ''], 'Витрати', 222);
  put(7, [D(c.addDays_(today, 4)), 40000, 'Зарплата', '', 'план', '', '', ''], 'Надходження', 333);
  put(8, [D(c.addDays_(today, -4)), 500, 'Таксі', '', 'план', '', '', ''], 'Витрати', 444);
  c.calendarOn();
  const cal = env.cals.find(x => x.name === 'Monitoring Budgets: платежі');
  const ev = cal.events.filter(e => !e.gone);
  assert.strictEqual(JSON.stringify(ev.map(e => e.title).sort()), JSON.stringify(['💳 Комуналка · 700 ₴', '💳 Підписки · 300 ₴']));
  assert.ok(ev.every(e => e.rem.length === 2 && e.getTag('mbs')));
  // план став фактом, а другий прибрали
  j.getRange('E5').setValue('✅ факт');
  j.getRange('A6:H6').setValues([['', '', '', '', '', '', '', '']]); j.getRange(6, 31).setValue('');
  c.calSync_(env.ss);
  const after = cal.events.filter(e => !e.gone);
  assert.strictEqual(JSON.stringify(after.map(e => e.title)), JSON.stringify(['✅ 💳 Комуналка · 700 ₴']));
  assert.strictEqual(after[0].rem.length, 0);
  assert.ok(!fe.gone && fe.title === '🔁 Оренда · 9 000 ₴', 'чужий календар не змінено');
});

t('never picks a calendar by name: only the one it created (id saved); turning off with «Так» deletes only that one', () => {
  const env = setup(), c = env.ctx;
  const foreign = env.ctx.CalendarApp.createCalendar('Monitoring Budgets: платежі', {});
  c.calendarOn();
  const mine = env.cals.filter(x => x.name === 'Monitoring Budgets: платежі');
  assert.strictEqual(mine.length, 2, 'створено власний, хоча з такою назвою вже був');
  c.calendarOff();
  assert.ok(!foreign.deleted && mine[1].deleted);
});

console.log(n + ' tests passed');
