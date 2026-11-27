// Тестовий помічник: таблиця, налаштована через панель «🚀 Підключення» (замість старого майстра з вікнами).
const vm = require('vm');

/**
 * o: { family: 'Олена, Іван', credit: true|false, bday: 25, fop: [['Олена', '3']], mods: ['tg', 'credit', …] (решта «ні»), demo: true }
 */
function panelSetup(env, o) {
  const c = env.ctx;
  o = o || {};
  env.ss.insertSheet('Аркуш1');
  c.startWizard();
  if (o.family) c.cnMode(true, o.family);
  c.cnCard(o.credit !== false);
  if (o.bday) c.cnBday(o.bday);
  (o.fop || []).forEach(f => c.cnFop(f[1], f[0]));
  if (o.mods) {
    const MODS = vm.runInContext('MODS', c), st = env.ss.getSheetByName('Налаштування');
    MODS.forEach((m, i) => st.getRange(44 + i, 2).setValue(o.mods.indexOf(m.id) >= 0 ? 'так' : 'ні'));
    c.modulesApply_(env.ss, true);
  }
  if (o.demo) c.cnDemo(true);
  c.cnFinish();
  return env;
}

module.exports = { panelSetup };
