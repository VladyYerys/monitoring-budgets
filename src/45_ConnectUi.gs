/** 45_ConnectUi.gs: HTML бічної панелі «🚀 Підключення» (дані й дії: 44_Connect.gs). */

/*
 * Один рядок HTML з CSS і JS усередині: так увесь код лишається в .gs-файлах (і «весь код одним файлом» працює).
 * Панель нічого не зберігає в браузері; токени з полів одразу йдуть у скрипт і в панель назад не повертаються.
 * Зовнішні адреси лише як посилання для людини (api.monobank.ua, t.me, gemini.google.com): панель їх не завантажує.
 */

function cnHtml_() {
  return String.raw`<!DOCTYPE html><html><head><base target="_blank"><meta charset="utf-8"><style>
body{font:13px/1.45 Arial,sans-serif;color:#202124;margin:0;padding:10px 12px 40px}
a{color:#1a73e8}button{font:13px Arial,sans-serif;padding:5px 10px;border:1px solid #dadce0;border-radius:4px;background:#fff;cursor:pointer;margin:2px 4px 2px 0}
button.pri{background:#1a73e8;color:#fff;border-color:#1a73e8}button:disabled{opacity:.5}
input,select,textarea{font:13px Arial,sans-serif;padding:4px 6px;border:1px solid #dadce0;border-radius:4px;box-sizing:border-box}
input[type=text],input[type=password]{width:100%}textarea{width:100%;height:60px;font:11px/1.35 monospace}
.mode{display:flex;gap:12px;margin-bottom:6px}.hide{display:none}.row{display:flex;gap:6px;align-items:center;margin:4px 0;flex-wrap:wrap}
.step{border:1px solid #dadce0;border-radius:8px;margin:8px 0}.sh{padding:9px 10px;cursor:pointer;display:flex;align-items:center;gap:8px;font-weight:bold}
.n{width:20px;height:20px;border-radius:50%;background:#e8eaed;display:inline-flex;align-items:center;justify-content:center;font-size:12px;flex:none}
.done .n{background:#188038;color:#fff}.st{font-weight:normal;color:#5f6368;font-size:12px;margin-left:auto;text-align:right}
.sb{padding:0 10px 10px;display:none}.open .sb{display:block}.t{font-weight:bold;margin:10px 0 4px}.hint{color:#5f6368;font-size:12px;margin:4px 0}
ol.how{padding-left:18px;margin:6px 0}.qr{text-align:center;margin:6px 0}.c{text-align:center;margin:4px 0}
.item{border:1px solid #e8eaed;border-radius:6px;padding:6px 8px;margin:4px 0}.item .sub{color:#5f6368;font-size:12px;margin-top:2px}
.ok{color:#188038}.warn{color:#b06000}.msg{margin:8px 0;min-height:18px}.foot{border-top:1px solid #e8eaed;padding-top:8px;margin-top:8px}
details{margin:8px 0}summary{cursor:pointer;color:#5f6368}
</style></head><body>
<div class="mode"><label><input type="radio" name="mode" id="m0" onchange="setMode(false)"> Я сам</label>
<label><input type="radio" name="mode" id="m1" onchange="showFam(true)"> Сім'я</label></div>
<div id="fam" class="hide"><div class="row"><input type="text" id="members" placeholder="Учасники через кому: Олена, Іван"></div>
<button onclick="setMode(true)">Зберегти учасників</button></div>

<div class="step" id="s1"><div class="sh" onclick="toggle(1)"><span class="n">1</span>Додайте картку<span class="st" id="st1"></span></div><div class="sb">
<p class="t">Mono</p>
<ol class="how"><li><a id="monoLink" href="#">Відкрийте api.monobank.ua</a>, відскануйте QR застосунком monobank і підтвердьте.</li>
<li>Скопіюйте токен сюди (він збережеться лише у вашому Google-акаунті):</li></ol>
<div class="row"><input type="password" id="mtok" autocomplete="off" placeholder="Токен Mono"></div>
<button onclick="monoCheck()">Перевірити</button>
<div id="mstat"></div><div id="maccs"></div><div id="mdone"></div>
<div id="fopq" class="hide"><p class="t">Рахунок ФОП</p><div class="row">Група: <select id="fopg"><option>1</option><option>2</option><option selected>3</option></select>
<select id="fopw" class="hide"></select><button onclick="setFop()">Зберегти</button></div><p class="hint" id="fopst"></p></div>
<p class="t">Інші банки</p>
<button onclick="call('connectPrivat')">ПриватБанк для ФОП</button><button onclick="call('importStatements')">Виписка з папки</button>
<p class="hint">Виписка: CSV чи Google Таблиця в папці «Monitoring Budgets: виписки» на Диску.</p>
<p class="t">Картка й дохід</p>
<div class="row">Картка: <select id="card" onchange="setCard()"><option value="1">кредитна з лімітом</option><option value="0">дебетова</option></select></div>
<div class="row">Дохід з <input type="number" id="bday" min="1" max="31" style="width:56px"> числа йде в наступний місяць <button onclick="setBday()">OK</button></div>
<label class="row"><input type="checkbox" id="sched" onchange="setSched()"> Імпорт з банку щодня (12:45, 18:45, 23:45)</label>
<button class="pri" onclick="go(2)">Далі</button></div></div>

<div class="step" id="s2"><div class="sh" onclick="toggle(2)"><span class="n">2</span>Сповіщення в Telegram<span class="st" id="st2"></span></div><div class="sb">
<ol class="how"><li>Створіть бота: відскануйте QR телефоном або <a id="bfLink" href="#">відкрийте @BotFather</a>. Напишіть /newbot, дайте назву й ім'я, що закінчується на bot.</li></ol>
<div class="qr" id="bfqr"></div>
<ol class="how" start="2"><li>Скопіюйте токен, який дасть BotFather:</li></ol>
<div class="row"><input type="password" id="ttok" autocomplete="off" placeholder="123456789:AAE…"></div>
<button onclick="tgToken()">Далі</button>
<div id="tlink" class="hide"><ol class="how" start="3"><li id="tl3">Відскануйте телефоном і натисніть <b>Start</b>:</li></ol>
<div class="qr" id="tqr"></div><p class="c"><a id="tbot" href="#">Відкрити бота</a></p><p class="c" id="tstat"></p>
<button id="tgrp" onclick="tgGroup()">QR для сімейної групи</button></div>
<div id="tdone"></div>
<button class="pri" onclick="go(3)">Далі</button></div></div>

<div class="step" id="s3"><div class="sh" onclick="toggle(3)"><span class="n">3</span>Регулярні платежі<span class="st" id="st3"></span></div><div class="sb">
<p class="hint">Оренда, підписки, зарплата стануть планами в журналі до кінця наступного місяця. Факт з банку (та сама категорія, ±5 днів, ±25 %) закриє план сам.</p>
<button onclick="scan(false)">Знайти в операціях</button>
<div id="sug"></div>
<details><summary>Додати свій</summary>
<div class="row"><input type="text" id="rn" placeholder="Назва: Оренда"></div>
<div class="row"><select id="rc" style="width:100%"></select></div>
<div class="row"><input type="number" id="ra" placeholder="Сума, ₴" style="width:90px"><input type="number" id="rd" placeholder="День" min="1" max="31" style="width:60px">
<select id="rp"><option>щомісяця</option><option>щокварталу</option><option>щороку</option></select></div>
<button onclick="addOwn()">Додати</button></details>
<div id="rlist"></div>
<button class="pri" onclick="go(4)">Далі</button></div></div>

<div class="step" id="s4"><div class="sh" onclick="toggle(4)"><span class="n">4</span>Передайте Gemini<span class="st" id="st4"></span></div><div class="sb">
<p class="hint">Знімок фінансів ляже на ваш Диск і оновлюватиметься щодня. Gemini читає його, лише коли ви питаєте. Коментарів, карток і токенів у ньому немає.</p>
<p class="t">Мої правила <span class="hint">(необов'язково; щодня йдуть у знімок, навичку перестворювати не треба)</span></p>
<div class="row"><label style="width:100%">Пріоритет №1<select id="ap" style="width:100%;margin-top:2px"></select></label></div>
<div class="row"><label>Подушка, % доходу <input type="number" id="ac" min="0" max="100" style="width:58px"></label>
<label>Запас на картці, ₴ <input type="number" id="af" min="0" style="width:84px"></label></div>
<div class="row"><label><input type="checkbox" id="ad"> Розстрочки й кредити: обов'язкові платежі</label></div>
<div class="row hide" id="afop"><label>ФОП: відкладати з доходу ФОП, % <input type="number" id="ao" min="0" max="100" style="width:58px"></label></div>
<textarea id="ar" style="height:54px;font:12px/1.4 Arial,sans-serif" placeholder="Ще мої правила, наприклад: на подарунки не більше 3 000 ₴ на місяць. Без номерів карток і паролів."></textarea>
<button class="pri" onclick="adv()">Зберегти й зробити текст для Gemini</button>
<div id="advres"></div></div></div>

<div id="msg" class="msg"></div>
<div class="foot"><button class="pri" onclick="finish()">Готово</button><button id="demoBtn" onclick="demo()">Демо-місяць</button></div>
<details><summary>Інше</summary>
<button onclick="call('importNow')">Імпорт з банку зараз</button><button onclick="call('disconnectMono')">Відключити Mono</button>
<button onclick="call('disconnectPrivat')">Відключити ПриватБанк</button><button onclick="call('disconnectTelegram')">Відключити Telegram</button>
<button onclick="call('calendarOff')">Календар: видалити</button><button onclick="call('protectionOnMenu')">Захист: увімкнути</button>
<button onclick="call('removeProtection')">Захист: прибрати</button>
<p class="hint">Модулі (так / ні): аркуш «Старт». Ця панель: меню «🚀 Підключення».</p></details>

<script>
  var S = null, monoToken = '', poll = null, polls = 0, polling = false, sug = [], tg = null, grp = false;
  function el(id) { return document.getElementById(id); }
  function h(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function money(x) { return Math.round(Number(x) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' ₴'; }
  function note(t, bad) { el('msg').innerHTML = t ? '<span class="' + (bad ? 'warn' : '') + '">' + t + '</span>' : ''; }
  function run(fn, args, ok, wait) {
    note(wait || '⏳ Хвилинку…');
    var r = google.script.run.withSuccessHandler(function (x) { note(''); if (ok) ok(x); })
      .withFailureHandler(function (e) { note('⚠️ ' + h(e && e.message || e), true); });
    r[fn].apply(r, args || []);
  }
  function call(fn) { run(fn, [], function () { load(); }); }
  function toggle(n) { var s = el('s' + n); s.className = s.className.indexOf('open') >= 0 ? s.className.replace(' open', '') : s.className + ' open'; }
  function go(n) { [1, 2, 3, 4].forEach(function (k) { var s = el('s' + k); s.className = s.className.replace(' open', '') + (k === n ? ' open' : ''); }); el('s' + n).scrollIntoView(); }
  function showFam(on) { el('fam').className = on ? '' : 'hide'; }
  function setMode(fam) { run('cnMode', [fam, el('members').value], render); }

  function render(s) {
    S = s;
    el('m0').checked = !s.family; el('m1').checked = s.family; showFam(s.family); el('members').value = s.members.join(', ');
    el('monoLink').href = s.links.mono; el('bfLink').href = s.links.botfather; el('bfqr').innerHTML = s.qrBotFather;
    var b = s.bank, done = [b.accounts.length > 0, s.tg.connected, s.reg.items.length > 0, s.adv.on && !!s.adv.doc];
    el('st1').innerHTML = b.accounts.length ? h(b.accounts.map(function (a) { return a.name; }).join(', ')) : 'не підключено';
    el('st2').innerHTML = s.tg.connected ? '@' + h(s.tg.bot) + (s.tg.on ? '' : ' (модуль вимкнено)') : s.tg.bot ? 'чекає Start' : 'не підключено';
    el('st3').innerHTML = s.reg.items.length ? s.reg.items.length + ' на аркуші' : 'ще немає';
    el('st4').innerHTML = !done[3] ? 'не підключено' : s.adv.old ? '<span class="warn">нова навичка: скопіюйте ще раз</span>' :
      s.adv.relink ? '<span class="warn">новий знімок: скопіюйте ще раз</span>' : s.adv.copied ? 'навичка v' + s.adv.copied : 'знімок є';
    advFill(s.adv);
    [1, 2, 3, 4].forEach(function (k) { var e = el('s' + k), open = e.className.indexOf('open') >= 0; e.className = 'step' + (done[k - 1] ? ' done' : '') + (open ? ' open' : ''); });
    el('card').value = b.credit ? '1' : '0'; el('bday').value = b.bday; el('sched').checked = b.schedule;
    el('sched').disabled = !(b.mono.length || b.privat.length);
    el('fopq').className = b.fopAcc ? '' : 'hide';
    el('fopst').innerHTML = b.fops.length ? '✓ ' + h(b.fops.join(', ')) : 'Вкажіть групу: податки й строки порахуються на аркуші «ФОП».';
    el('fopw').className = s.family ? '' : 'hide';
    el('fopw').innerHTML = s.members.map(function (m) { return '<option>' + h(m) + '</option>'; }).join('');
    el('mstat').innerHTML = (b.mono.length ? '<p class="ok">✓ Mono: ' + h(b.mono.join(', ')) + '</p>' : '') +
      (b.privat.length ? '<p class="ok">✓ ПриватБанк: ' + h(b.privat.join(', ')) + '</p>' : '');
    el('tdone').innerHTML = s.tg.connected ? '<p class="ok">✓ Бот @' + h(s.tg.bot) + ' підключено. Підсумок щодня о 05:00.</p>' : '';
    if (s.tg.link && !poll) { tg = s.tg.link; showTg(); startPoll(); }
    el('rc').innerHTML = catOptions('');
    el('rlist').innerHTML = s.reg.items.length ? '<p class="t">На аркуші «Регулярні»</p>' + s.reg.items.map(function (r) {
      return '<div class="item">' + h(r.name) + '<div class="sub">' + money(r.amount) + ' · ' + r.day + '-го · ' + h(r.per) + (r.on ? '' : ' · вимкнено') +
        (r.state ? '<br>' + h(r.state) : '') + '</div></div>'; }).join('') : '';
    el('demoBtn').textContent = s.demo ? 'Прибрати демо' : 'Демо-місяць';
    el('demoBtn').style.display = s.demo || s.empty ? '' : 'none';
    if (!document.querySelector('.step.open')) { var first = done.indexOf(false); if (first >= 0 && !s.done) go(first + 1); }
  }
  function load() { run('cnState', [], render, ' '); }

  function monoCheck() {
    monoToken = el('mtok').value.trim();
    if (!monoToken) { note('Вставте токен Mono.', true); return; }
    run('cnMonoCheck', [monoToken], function (list) {
      var who = S.family ? '<div class="row">Чиї: <select id="mwho">' + S.who.map(function (w) { return '<option>' + h(w) + '</option>'; }).join('') + '</select></div>' : '';
      el('maccs').innerHTML = '<p class="hint">Які рахунки вести?</p>' + list.map(function (a) {
        return '<label class="item" style="display:block"><input type="checkbox" class="macc" value="' + h(a.id) + '"' + (a.pick ? ' checked' : '') + '> ' + h(a.label) +
          '<div class="sub">' + (a.fop ? 'рахунок ФОП' : a.limit ? 'кредитна, ліміт ' + money(a.limit) : 'без ліміту') + '</div></label>'; }).join('') + who +
        '<button class="pri" onclick="monoSave()">Зберегти й завантажити операції</button>';
    }, '⏳ Питаю Mono…');
  }
  function monoSave() {
    var ids = Array.prototype.slice.call(document.querySelectorAll('.macc:checked')).map(function (x) { return x.value; });
    if (!ids.length) { note('Позначте хоча б один рахунок.', true); return; }
    run('cnMonoSave', [monoToken, ids, el('mwho') ? el('mwho').value : ''], function (r) {
      monoToken = ''; el('mtok').value = ''; el('maccs').innerHTML = '';
      var lines = ['✓ Збережено: ' + h(r.saved.join(', ')) + '.'];
      if (r.credit !== null) lines.push(r.credit ? 'Картка кредитна: «мої» рахуються окремо від ліміту (модуль «Ліміт» увімкнено).' : 'Картка без ліміту: модуль «Ліміт» вимкнено.');
      if (r.fop) lines.push('Є рахунок ФОП: модуль «ФОП» увімкнено, вкажіть групу нижче.');
      lines.push('Імпорт за розкладом увімкнено.');
      el('mdone').innerHTML = '<p class="ok">' + lines.join('<br>') + '</p>';
      run('cnImport', [], function (x) {
        el('mdone').innerHTML += '<p class="ok">✓ Перший імпорт: ' + h(x.text) + (x.own != null ? '. Зараз на картці ' + money(x.own) : '') + '.</p>';
        load();
      }, '⏳ Завантажую операції (до 1–2 хв)…');
    }, '⏳ Зберігаю…');
  }
  function setCard() { run('cnCard', [el('card').value === '1'], render); }
  function setBday() { run('cnBday', [Number(el('bday').value)], render); }
  function setSched() { run('cnSchedule', [el('sched').checked], render); }
  function setFop() { run('cnFop', [el('fopg').value, S.family ? el('fopw').value : ''], render); }

  function showTg() {
    el('tlink').className = '';
    el('tqr').innerHTML = grp ? tg.qrGroup : tg.qr; el('tbot').href = grp ? tg.group : tg.private;
    el('tbot').textContent = grp ? 'Додати бота в групу' : 'Відкрити @' + tg.bot;
    el('tl3').innerHTML = grp ? 'Відскануйте телефоном, виберіть сімейну групу й додайте бота:' : 'Відскануйте телефоном і натисніть <b>Start</b>:';
    el('tgrp').textContent = grp ? 'QR для особистого чату' : 'QR для сімейної групи';
  }
  function tgToken() {
    var t = el('ttok').value.trim();
    if (!t) { note('Вставте токен бота.', true); return; }
    run('cnTgToken', [t], function (r) { el('ttok').value = ''; tg = r; grp = false; showTg(); startPoll(); load(); }, '⏳ Перевіряю бота…');
  }
  function tgGroup() { grp = !grp; showTg(); }
  function stopPoll() { if (poll) clearInterval(poll); poll = null; }
  function startPoll() {
    stopPoll(); polls = 0; el('tstat').textContent = 'Чекаю на Start…';
    poll = setInterval(function () {
      if (polling) return;
      if (++polls > 100) { stopPoll(); el('tstat').innerHTML = 'Не бачу Start. <a href="#" onclick="startPoll();return false">Перевірити ще раз</a>'; return; }
      polling = true;
      google.script.run.withSuccessHandler(function (r) {
        polling = false;
        if (r && r.ok) { stopPoll(); el('tlink').className = 'hide'; el('tdone').innerHTML = '<p class="ok">✓ Підключено: ' + h(r.chat) + '. Гляньте Telegram: там перше повідомлення.</p>'; load(); }
      }).withFailureHandler(function (e) { polling = false; stopPoll(); el('tstat').textContent = '⚠️ ' + (e && e.message || e); }).cnTgPoll();
    }, 3000);
  }

  function catOptions(sel) {
    return '<option value="">категорія…</option>' + (S ? S.cats : []).map(function (c) {
      return '<option' + (c.name === sel ? ' selected' : '') + '>' + h(c.name) + '</option>'; }).join('');
  }
  function scan(deep) {
    run('cnRegSuggest', [deep], function (r) {
      sug = r.items;
      var html = sug.length ? '<p class="hint">Знайшли за ' + r.days + ' днів. Позначте потрібні:</p>' + sug.map(function (x, i) {
        return '<div class="item"><label><input type="checkbox" class="sg" data-i="' + i + '" checked> ' + h(x.name) + '</label><div class="sub">' +
          (x.income ? 'надходження ' : '') + money(x.amount) + ' · ' + x.day + '-го · ' + x.months + ' міс.</div><select class="sgc" data-i="' + i + '" style="width:100%">' +
          catOptions(x.cat) + '</select></div>'; }).join('') + '<button class="pri" onclick="addSel()">Додати вибрані</button>'
        : '<p class="hint">Повторюваних платежів не знайшли' + (r.days ? ' (дані за ' + r.days + ' днів)' : '') + '. Додайте свій нижче.</p>';
      if (r.income && S.bank.bday !== r.income.day) html += '<p class="hint">Схоже, дохід «' + h(r.income.name) + '» приходить ≈ ' + r.income.day + '-го. ' +
        '<a href="#" onclick="useBday(' + r.income.day + ');return false">Дохід з ' + r.income.day + '-го йде в наступний місяць</a></p>';
      if (r.canDeep) html += '<button onclick="scan(true)">Шукати глибше: виписка Mono за 2 місяці (1–2 хв)</button>';
      el('sug').innerHTML = html;
    }, deep ? '⏳ Дивлюсь виписку Mono за 2 місяці (1–2 хв, Mono дає місяць за запит раз на хвилину)…' : '⏳ Шукаю…');
  }
  function useBday(d) { el('bday').value = d; setBday(); }
  function addSel() {
    var list = [];
    Array.prototype.slice.call(document.querySelectorAll('.sg:checked')).forEach(function (x) {
      var i = Number(x.getAttribute('data-i')), c = document.querySelector('.sgc[data-i="' + i + '"]').value;
      list.push({ name: sug[i].name, cat: c, amount: sug[i].amount, day: sug[i].day, per: 'щомісяця' });
    });
    if (!list.length) { note('Позначте хоча б один.', true); return; }
    if (list.some(function (x) { return !x.cat; })) { note('Виберіть категорію для кожного позначеного.', true); return; }
    addRules(list);
  }
  function addOwn() {
    var x = { name: el('rn').value.trim(), cat: el('rc').value, amount: Number(el('ra').value), day: Number(el('rd').value), per: el('rp').value };
    if (!x.name || !x.cat || !(x.amount > 0) || !(x.day >= 1 && x.day <= 31)) { note('Заповніть назву, категорію, суму й день (1–31).', true); return; }
    addRules([x]);
  }
  function addRules(list) {
    run('cnRegAdd', [list], function (r) {
      el('sug').innerHTML = ''; ['rn', 'ra', 'rd'].forEach(function (k) { el(k).value = ''; });
      note('✓ Додано: ' + r.added + ', планів у журналі: +' + r.plans + (r.skipped.length ? '. Пропущено: ' + h(r.skipped.join(', ')) : '') +
        (r.full ? '. На аркуші закінчились рядки.' : ''));
      render(r.state);
    }, '⏳ Додаю…');
  }

  function area(t) { return '<textarea readonly onclick="this.select()">' + h(t) + '</textarea>'; }
  var advInit = false;
  function advFill(a) {
    el('afop').className = 'row' + (a.fop ? '' : ' hide');
    if (advInit) return;
    advInit = true;
    el('ap').innerHTML = a.prio.map(function (p) { return '<option value="' + h(p) + '">' + h(p) + '</option>'; }).join('');
    var r = a.rules;
    el('ap').value = r.prio; el('ac').value = r.cushion ? Math.round(r.cushion * 1000) / 10 : ''; el('af').value = r.floor || '';
    el('ad').checked = r.debtsMust; el('ao').value = r.fopPct ? Math.round(r.fopPct * 1000) / 10 : ''; el('ar').value = r.rules; el('ar').maxLength = a.max;
  }
  function advRules() {
    return { prio: el('ap').value, cushion: Number(el('ac').value) || 0, floor: Number(el('af').value) || 0, debtsMust: el('ad').checked,
      fopPct: Number(el('ao').value) || 0, rules: el('ar').value };
  }
  function adv() {
    run('cnAdvisor', [advRules()], function (r) {
      el('advres').innerHTML = '<p class="ok">✓ Правила збережено. Знімок: <a href="' + h(r.doc) + '">' + h(r.docName) + '</a></p><ol class="how">' +
        '<li><button class="pri" onclick="cpMsg()">📋 Скопіювати для Gemini</button> <span id="cpok"></span></li>' +
        '<li><a href="' + h(r.gemini) + '">Відкрийте Gemini</a>, вставте в чат і надішліть: Gemini сам створить навичку «' + h(r.title) + '». ' +
        'У тексті пряме посилання на ваш знімок: Gemini відкриє його одразу, без пошуку.</li>' +
        '<li>Далі просто питайте в будь-якому чаті Gemini, наприклад: «' + h(r.starters[0]) + '». Google Диск у Gemini має бути підключений.</li></ol>' +
        '<textarea id="advmsg" readonly onclick="this.select()">' + h(r.msg) + '</textarea>' +
        '<details><summary>Інакше: створити вручну або з файла</summary><p class="hint">Gemini → Skills → Create → Create manually (поля нижче, з посиланням) або Upload <a href="' +
        h(r.skillFile) + '">SKILL.md</a> (без посилання: Gemini шукатиме знімок за назвою).</p>Назва:' + area(r.title) + 'Опис:' + area(r.description) + 'Інструкція:' + area(r.instructions) + '</details>' +
        '<p class="hint">Правила можна змінити тут будь-коли: вони йдуть у знімок, а навичка (версія ' + r.v + ') лишається та сама.</p>';
      load();
    }, '⏳ Зберігаю й роблю знімок…');
  }
  function cpMsg() {
    var t = el('advmsg');
    function ok() { el('cpok').innerHTML = '<span class="ok">✓ Скопійовано</span>'; google.script.run.withSuccessHandler(load).advSkillCopied(); }
    function fb() { t.select(); var d = false; try { d = document.execCommand('copy'); } catch (e) { d = false; } if (d) ok(); else el('cpok').innerHTML = 'Виділено: натисніть Cmd+C або Ctrl+C'; }
    try { navigator.clipboard.writeText(t.value).then(ok, fb); } catch (e) { fb(); }
  }
  function demo() { run('cnDemo', [!S.demo], render, S.demo ? '⏳ Прибираю демо…' : '⏳ Додаю демо-місяць…'); }
  function finish() {
    run('cnFinish', [], function (r) {
      note('<span class="ok">✓ Готово. Чекліст «Старту»: ' + r.done + ' з ' + r.need + '.</span>' + (r.warnings.length ? '<br>' + h(r.warnings.join(' ')) : '') +
        ' Модулі: аркуш «Старт». <a href="#" onclick="google.script.host.close();return false">Закрити</a>');
    });
  }
  load();
</script></body></html>`;
}
