// Minimal Google Apps Script mock (Sheets grid, Drive, triggers, props, UrlFetch) for logic tests.
const vm = require('vm'), fs = require('fs');

function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function colStr(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function parseA1(a1) {
  const m = a1.replace(/\$/g, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!m) throw new Error('bad a1 ' + a1);
  const c0 = colNum(m[1]), r0 = +m[2], c1 = m[3] ? colNum(m[3]) : c0, r1 = m[4] ? +m[4] : r0;
  return { r: r0, c: c0, nr: r1 - r0 + 1, nc: c1 - c0 + 1 };
}

function makeEnv(opts = {}) {
  const tz = opts.tz || 'Europe/Kiev';
  const calls = { sleep: 0, fetch: [], dialogs: [], alerts: [], prompts: [], logs: [], textFinder: [], toasts: [] };
  // ---------- Sheets ----------
  let protId = 0;
  function Protection(sheet, type) { this.id = ++protId; this.sheet = sheet; this.type = type; this.desc = ''; this.warn = false; this.unprot = []; }
  Protection.prototype.getDescription = function () { return this.desc; };
  Protection.prototype.setDescription = function (d) { this.desc = d; return this; };
  Protection.prototype.setWarningOnly = function (w) { this.warn = w; return this; };
  Protection.prototype.setUnprotectedRanges = function (r) { this.unprot = r.map(x => x.getA1Notation()); return this; };
  Protection.prototype.remove = function () { this.sheet.prots = this.sheet.prots.filter(p => p !== this); };

  function DV(spec) { this.spec = spec; }
  DV.prototype.copy = function () { const b = new DVB(Object.assign({}, this.spec)); if (this.help) b.help = this.help; if (this.reject) b.reject = true; return b; };
  function DVB(spec) { this.spec = spec || {}; }
  // reject: setAllowInvalid(false), як у Таблицях: запис значення не зі списку падає з помилкою
  DVB.prototype.requireValueInList = function (l) { this.spec = { kind: 'list', list: l.slice() }; return this; };
  DVB.prototype.requireValueInRange = function (r) { this.spec = { kind: 'range', range: r.getA1Notation(), sheet: r.getSheet().getName() }; return this; };
  DVB.prototype.setAllowInvalid = function (a) { this.reject = a === false; return this; };
  DVB.prototype.requireDate = function () { this.spec = { kind: 'date' }; return this; };
  DVB.prototype.requireNumberBetween = function (a, b) { this.spec = { kind: 'between', a: a, b: b }; return this; };
  DVB.prototype.requireNumberGreaterThan = function (a) { this.spec = { kind: 'gt', a: a }; return this; };
  DVB.prototype.setHelpText = function (t) { this.help = t; return this; };
  DV.prototype.getCriteriaValues = function () { return this.spec.kind === 'list' ? [this.spec.list.slice(), true] : []; };
  DV.prototype.getHelpText = function () { return this.help || ''; };
  DV.prototype.getCriteriaType = function () { return { list: 'VALUE_IN_LIST', range: 'VALUE_IN_RANGE', date: 'DATE_IS_VALID_DATE', between: 'NUMBER_BETWEEN', gt: 'NUMBER_GREATER_THAN' }[this.spec.kind] || 'CUSTOM_FORMULA'; };
  DV.prototype.getAllowInvalid = function () { return !this.reject; };
  DVB.prototype.build = function () { const d = new DV(this.spec); if (this.help) d.help = this.help; if (this.reject) d.reject = true; return d; };

  function CFR(spec) { this.spec = spec; }
  CFR.prototype.getBooleanCondition = function () { const s = this.spec; const v = s.formula || s.startsWith || s.contains; return v ? { getCriteriaValues: () => [v] } : null; };
  CFR.prototype.getRanges = function () { return this.spec.ranges; };
  CFR.prototype.copy = function () { return new CFB(Object.assign({}, this.spec)); };
  function CFB(spec) { this.spec = spec || {}; }
  ['setBackground', 'setFontColor', 'setBold', 'setItalic', 'whenNumberLessThan', 'whenNumberGreaterThan', 'whenNumberEqualTo'].forEach(k => { CFB.prototype[k] = function (v) { this.spec[k] = v; return this; }; });
  CFB.prototype.whenFormulaSatisfied = function (f) { this.spec.formula = f; return this; };
  CFB.prototype.whenTextStartsWith = function (t) { this.spec.startsWith = t; return this; };
  CFB.prototype.whenTextContains = function (t) { this.spec.contains = t; return this; };
  CFB.prototype.setRanges = function (r) { this.spec.ranges = r; return this; };
  CFB.prototype.build = function () { return new CFR(this.spec); };

  function Sheet(ss, name) {
    this.ss = ss; this.name = name; this.cells = new Map(); this.prots = []; this.cf = []; this.hidden = [];
    this.dv = new Map(); this.formats = []; this.textFmt = new Set();
    this.maxR = 1000; this.maxC = 26;                          // нова вкладка в Таблицях: 1000 рядків × 26 колонок (A…Z)
  }
  const S = Sheet.prototype;
  S.key = (r, c) => r + ',' + c;
  S.cell = function (r, c) { return this.cells.get(this.key(r, c)) || { v: '', f: '' }; };
  S.put = function (r, c, v, f) { this.cells.set(this.key(r, c), { v: v, f: f || '' }); };
  S.getName = function () { return this.name; };
  S.getParent = function () { return this.ss; };
  S.getRange = function (a, b, nr, nc) {
    const g = typeof a === 'string' ? parseA1(a) : { r: a, c: b, nr: nr || 1, nc: nc || 1 };
    // як у Таблицях: діапазон за межами аркуша → помилка (рядки й колонки треба спершу додати)
    if (g.r + g.nr - 1 > this.maxR || g.c + g.nc - 1 > this.maxC) throw new Error('The coordinates of the range are outside the dimensions of the sheet. (' + this.name + ' ' +
      colStr(g.c) + g.r + ':' + colStr(g.c + g.nc - 1) + (g.r + g.nr - 1) + ', аркуш ' + this.maxR + '×' + this.maxC + ')');
    return new Range(this, g.r, g.c, g.nr, g.nc);
  };
  S.insertRowsAfter = function (r, n) { this.maxR += n; return this; };
  S.getMaxRows = function () { return this.maxR; };
  S.getLastRow = function () { let m = 0; for (const [k, x] of this.cells) if (x.v !== '' && x.v !== null || x.f) m = Math.max(m, +k.split(',')[0]); return m; };
  S.getProtections = function (t) { return this.prots.filter(p => p.type === t); };
  S.protect = function () { const p = new Protection(this, 'SHEET'); this.prots.push(p); return p; };
  S.getConditionalFormatRules = function () { return this.cf.slice(); };
  S.setConditionalFormatRules = function (r) { this.cf = r.slice(); };
  S.hideColumns = function (c, n) { this.hidden.push([c, n]); };
  S.showColumns = function (c, n) { this.hidden = this.hidden.filter(h => !(h[0] === c && h[1] === n)); this.shownCols = (this.shownCols || []).concat([[c, n]]); };
  S.isColumnHiddenByUser = function (c) { return this.hidden.some(h => c >= h[0] && c < h[0] + h[1]); };
  S.setFrozenRows = function () { }; S.setFrozenColumns = function () { }; S.setColumnWidth = function () { }; S.setRowHeight = function () { };
  S.setColumnWidths = function () { return this; }; S.setRowHeights = function () { return this; };
  S.setHiddenGridlines = function (v) { this.noGrid = !!v; return this; };
  S.deleteRows = function (r, n) { for (const k of Array.from(this.cells.keys())) if (+k.split(',')[0] >= r) this.cells.delete(k); this.maxR -= n; };
  S.deleteColumns = function (c, n) { for (const k of Array.from(this.cells.keys())) if (+k.split(',')[1] >= c) this.cells.delete(k); this.maxC -= n; };
  S.hideSheet = function () { this.hiddenSheet = true; return this; };
  S.showSheet = function () { this.hiddenSheet = false; return this; };
  S.isSheetHidden = function () { return !!this.hiddenSheet; };
  S.setName = function (n) { this.name = n; return this; };
  S.hideRows = function (r, n) { this.hiddenRows = (this.hiddenRows || []).concat([[r, n || 1]]); };
  S.showRows = function (r, n) { const k = n || 1; this.hiddenRows = (this.hiddenRows || []).filter(h => h[0] + h[1] <= r || h[0] >= r + k); };
  S.isRowHiddenByUser = function (r) { return (this.hiddenRows || []).some(h => r >= h[0] && r < h[0] + h[1]); };
  S.setTabColor = function (c) { this.tab = c; return this; };
  S.getMaxColumns = function () { return this.maxC; };
  S.clear = function () { this.cells = new Map(); this.dv = new Map(); return this; };
  S.getLastColumn = function () { let m = 0; for (const [k, x] of this.cells) if (x.v !== '' && x.v !== null || x.f) m = Math.max(m, +k.split(',')[1]); return m; };
  S.getActiveRange = function () { return this.getRange(this.activeRow || 1, 1); };
  S.createTextFinder = function (t) { return new TF([this], t); };
  S.insertColumnsAfter = function (c, n) { calls.sheetOps = calls.sheetOps || []; calls.sheetOps.push([this.name, 'insertColumnsAfter', c, n]); this.maxC += n; };
  S.insertColumnsBefore = function (c, n) { calls.sheetOps = calls.sheetOps || []; calls.sheetOps.push([this.name, 'insertColumnsBefore', c, n]); };
  S.deleteColumn = function (c) { calls.sheetOps = calls.sheetOps || []; calls.sheetOps.push([this.name, 'deleteColumn', c]); };
  S.getDataRange = function () { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); };
  S.insertRowsBefore = function (r0, n) {
    calls.sheetOps = calls.sheetOps || []; calls.sheetOps.push([this.name, 'insertRowsBefore', r0, n]); this.maxR += n;
    const shift = m => { const o = new Map(); for (const [k, x] of m) { const [r, c] = k.split(',').map(Number); o.set(this.key(r >= r0 ? r + n : r, c), x); } return o; };
    this.cells = shift(this.cells); this.dv = shift(this.dv);
    (this.charts || []).forEach(ch => { if (ch.row >= r0) ch.row += n; });
  };
  S.getSheetId = function () { return this.gid || (this.gid = 1000 + this.ss.getSheets().indexOf(this)); };
  S.getCharts = function () {
    return (this.charts || []).map(ch => ({ _o: ch, getContainerInfo: () => ({ getAnchorRow: () => ch.row, getAnchorColumn: () => ch.col || 1, getOffsetY: () => 0 }),
      getOptions: () => ({ get: k => (k === 'height' ? ch.h : (ch.opts || {})[k] === undefined ? null : ch.opts[k]) }) }));
  };
  S.newChart = function () {
    const spec = { ranges: [], opts: {} }, sh = this;
    const b = { setChartType: t => (spec.type = t, b), addRange: r => (spec.ranges.push(r.getSheet().getName() + '!' + r.getA1Notation()), b),
      setNumHeaders: n => (spec.headers = n, b), setPosition: (r, c) => (spec.row = r, spec.col = c, b), setOption: (k, v) => (spec.opts[k] = v, b),
      build: () => ({ spec: spec, sheet: sh }) };
    return b;
  };
  S.insertChart = function (c) { this.charts = this.charts || []; this.charts.push({ row: c.spec.row, col: c.spec.col, h: c.spec.opts.height, opts: c.spec.opts, ranges: c.spec.ranges, type: c.spec.type }); };
  S.removeChart = function (w) { this.charts = (this.charts || []).filter(ch => ch !== w._o); };

  function Range(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  const R = Range.prototype;
  R.grid = function (fn) { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(fn(this.r + i, this.c + j)); o.push(row); } return o; };
  R.getValues = function () { return this.grid((r, c) => this.sh.cell(r, c).v); };
  R.getValue = function () { return this.sh.cell(this.r, this.c).v; };
  R.getDisplayValue = function () { const v = this.getValue(); return v instanceof Date ? v.toISOString() : String(v); };
  R.getFormulas = function () { return this.grid((r, c) => this.sh.cell(r, c).f); };
  R.getFormula = function () { return this.sh.cell(this.r, this.c).f; };
  R.setValues = function (vals) {
    if (vals.length !== this.nr || vals.some(r => r.length !== this.nc)) throw new Error('setValues size ' + vals.length + 'x' + vals[0].length + ' vs ' + this.nr + 'x' + this.nc + ' at ' + this.getA1Notation());
    vals.forEach((row, i) => row.forEach((v, j) => {                         // перевірка даних з відмовою: увесь запис падає
      const d = this.sh.dv.get(this.sh.key(this.r + i, this.c + j));
      if (!d || !d.reject || v === '' || v === null || v === undefined || (typeof v === 'string' && v.charAt(0) === '=')) return;
      const list = d.spec.kind === 'list' ? d.spec.list.map(String) : null;    // список із діапазону не перевіряємо: формули мок не рахує
      if (list && list.indexOf(String(v)) < 0) throw new Error('The data you entered in cell ' + this.sh.getRange(this.r + i, this.c + j).getA1Notation() +
        ' violates the data validation rules set on this cell. Please enter one of the following values: ' + list.join(', ') + '.');
    }));
    vals.forEach((row, i) => row.forEach((v, j) => {
      if (typeof v === 'string' && v.charAt(0) === '=') { calls.valueFormulas = (calls.valueFormulas || []).concat([this.sh.name + '!' + this.sh.getRange(this.r + i, this.c + j).getA1Notation()]);
        this.sh.put(this.r + i, this.c + j, '', v); }
      else this.sh.put(this.r + i, this.c + j, sheetsValue(this.sh, this.r + i, this.c + j, v === undefined ? '' : v), '');
    })); return this;
  };
  R.setValue = function (v) { return this.setValues([[v]]); };
  // проба роздільника (02_Helpers.gs → fsepProbe_): =SUM(1,2) дає 3 лише в таблиці з «,» між аргументами (opts.sep, типово «,»)
  const probe = f => (f === '=SUM(1,2)' ? (opts.sep === ';' ? 1.2 : 3) : f === '=SUM(1;2)' ? (opts.sep === ';' ? 3 : '#ERROR!') : undefined);
  R.setFormulas = function (fs) {
    if (fs.length !== this.nr || fs.some(r => r.length !== this.nc)) throw new Error('setFormulas size ' + fs.length + 'x' + (fs[0] || []).length + ' vs ' + this.nr + 'x' + this.nc + ' at ' + this.getA1Notation());
    fs.forEach((row, i) => row.forEach((f, j) => { const x = this.sh.cell(this.r + i, this.c + j); const p = probe(f);
    this.sh.put(this.r + i, this.c + j, p === undefined ? x.v : p, f); })); return this; };
  R.setFormula = function (f) { return this.setFormulas([[f]]); };
  R.setNote = function (t) { this.sh.notes = this.sh.notes || {}; this.sh.notes[this.getA1Notation()] = t; return this; };
  R.getNotes = function () { const n = this.sh.notes || {}; return Array.from({ length: this.nr }, (x, i) => Array.from({ length: this.nc }, (y, j) =>
    n[this.sh.getRange(this.r + i, this.c + j).getA1Notation()] || '')); };
  R.setNotes = function (a) { a.forEach((row, i) => row.forEach((t, j) => this.sh.getRange(this.r + i, this.c + j).setNote(t))); return this; };
  R.getDataValidations = function () { return Array.from({ length: this.nr }, (x, i) => Array.from({ length: this.nc }, (y, j) =>
    this.sh.dv.get(this.sh.key(this.r + i, this.c + j)) || null)); };
  R.createTextFinder = function (t) { const rg = this; const tf = new TF([rg.sh], t); tf.range = rg.getA1Notation(); calls.textFinder.push(['range', rg.sh.name, tf.range, t]); return tf; };
  R.getDataValidation = function () { return this.sh.dv.get(this.sh.key(this.r, this.c)) || null; };
  R.setDataValidation = function (d) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) { if (d) this.sh.dv.set(this.sh.key(this.r + i, this.c + j), d); else this.sh.dv.delete(this.sh.key(this.r + i, this.c + j)); } return this; };
  R.setDataValidations = function (a) { a.forEach((row, i) => row.forEach((d, j) => this.sh.getRange(this.r + i, this.c + j).setDataValidation(d))); return this; };
  R.clearDataValidations = function () { return this.setDataValidation(null); };
  R.offset = function (r, c, nr, nc) { return new Range(this.sh, this.r + r, this.c + c, nr || this.nr, nc || this.nc); };
  R.getCell = function (r, c) { return new Range(this.sh, this.r + r - 1, this.c + c - 1, 1, 1); };
  R.getA1Notation = function () { return colStr(this.c) + this.r + (this.nr > 1 || this.nc > 1 ? ':' + colStr(this.c + this.nc - 1) + (this.r + this.nr - 1) : ''); };
  R.getSheet = function () { return this.sh; };
  R.getRow = function () { return this.r; }; R.getLastRow = function () { return this.r + this.nr - 1; };
  R.getNumRows = function () { return this.nr; }; R.getNumColumns = function () { return this.nc; };
  R.getColumn = function () { return this.c; }; R.getLastColumn = function () { return this.c + this.nc - 1; };
  R.copyTo = function (dst, type) {
    this.sh.formats.push([this.getA1Notation(), dst.getA1Notation(), type]);
    if (type === 'PASTE_DATA_VALIDATION') { const d = this.getDataValidation(); if (d) dst.setDataValidation(d); }
  };
  R.clearContent = function () { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.cells.delete(this.sh.key(this.r + i, this.c + j)); return this; };
  R.mergeAcross = function () { for (let i = 0; i < this.nr; i++) this.sh.getRange(this.r + i, this.c, 1, this.nc).merge(); return this; };
  R.merge = function () { this.sh.merges = this.sh.merges || []; if (this.sh.merges.indexOf(this.getA1Notation()) < 0) this.sh.merges.push(this.getA1Notation()); return this; };
  R.setNumberFormat = function (f) {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) {
      const k = this.sh.key(this.r + i, this.c + j); if (f === '@') this.sh.textFmt.add(k); else this.sh.textFmt.delete(k);
    } return this;
  };
  // як справжні Таблиці (opts.sheetsLike): «'текст» → текст; «2026-11» / «2026-11-05» у клітинці без формату «@» → дата
  function sheetsValue(sh, r, c, v) {
    if (!(opts.sheetsLike || process.env.SHEETS_LIKE) || typeof v !== 'string') return v;
    if (v.charAt(0) === "'") return v.slice(1);
    if (sh.textFmt.has(sh.key(r, c))) return v;
    const m = v.match(/^(\d{4})-(\d{2})(?:-(\d{2}))?$/);
    return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +(m[3] || 1), 12)) : v;
  }
  R.getFontFamily = function () { return (this.sh.fonts || {})[this.getA1Notation()] || 'Arial'; };
  R.setFontFamily = function (f) { this.sh.fonts = this.sh.fonts || {}; this.sh.fonts[this.getA1Notation()] = f; return this; };
  R.breakApart = function () { const a = this; this.sh.merges = (this.sh.merges || []).filter(m => { const g = parseA1(m);
    return g.r + g.nr - 1 < a.r || g.r > a.r + a.nr - 1 || g.c + g.nc - 1 < a.c || g.c > a.c + a.nc - 1; }); return this; };
  ['setFontWeight', 'setBackground', 'setFontSize', 'setWrap', 'setVerticalAlignment', 'setHorizontalAlignment', 'setFontColor', 'setFontStyle', 'setBorder'].forEach(k => { R[k] = function () { return this; }; });
  R.getFontSize = function () { return 11; }; R.getFontWeight = function () { return 'bold'; };
  R.getFontColor = function () { return '#3c4043'; }; R.getBackground = function () { return '#eef2fb'; };

  function TF(sheets, text) { this.sheets = sheets; this.text = text; }
  TF.prototype.matchCase = function () { return this; };
  TF.prototype.matchFormulaText = function () { return this; };
  TF.prototype.findAll = function () { const o = []; this.sheets.forEach(sh => { for (const [, x] of sh.cells) if (x.f && x.f.indexOf(this.text) >= 0) o.push(x); }); return o; };
  TF.prototype.replaceAllWith = function (to) {
    let n = 0; calls.textFinder.push([this.text, to]);
    this.sheets.forEach(sh => { for (const [, x] of sh.cells) if (x.f && x.f.indexOf(this.text) >= 0) { n += x.f.split(this.text).length - 1; x.f = x.f.split(this.text).join(to); } });
    return n;
  };

  const sheets = [];
  const ss = {
    getSheets: () => sheets, getSheetByName: n => sheets.find(s => s.name === n) || null,
    deleteSheet: sh => { const i = sheets.indexOf(sh); if (i >= 0) sheets.splice(i, 1); },
    insertSheet: (n, idx) => { const s = new Sheet(ss, n); sheets.splice(Math.min(idx === undefined ? sheets.length : idx, sheets.length), 0, s); return s; },
    getSpreadsheetTimeZone: () => tz, getId: () => 'SS1', getName: () => 'Test', getUrl: () => 'https://docs.google.com/spreadsheets/d/SS1/edit',
    getActiveSheet: () => ss.active || sheets[0], setActive: sh => { ss.active = sh; }, setActiveSheet: sh => { ss.active = sh; return sh; },
    moveActiveSheet: pos => { const sh = ss.active || sheets[0], i = sheets.indexOf(sh); sheets.splice(i, 1); sheets.splice(pos - 1, 0, sh); },
    createTextFinder: t => new TF(sheets, t),
    toast: (m, t) => { calls.toasts.push([t, m]); },
    // іменовані діапазони: ім'я → {sh, r, c, nr, nc}; переносом блоку в тестах керує move(name, a1)
    named: new Map(),
    getRangeByName: n => { const x = ss.named.get(n); return x ? new Range(x.sh, x.r, x.c, x.nr, x.nc) : null; },
    setNamedRange: (n, r) => { if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(n)) throw new Error('bad name ' + n); ss.named.set(n, { sh: r.sh, r: r.r, c: r.c, nr: r.nr, nc: r.nc }); },
    removeNamedRange: n => { ss.named.delete(n); },
    getNamedRanges: () => Array.from(ss.named.keys()).map(n => ({ getName: () => n, getRange: () => ss.getRangeByName(n), remove: () => ss.named.delete(n),
      setRange: r => ss.setNamedRange(n, r) })),
  };
  // ---------- Drive (for backups) ----------
  let fid = 0, clock = Date.UTC(2026, 10, 19);
  const files = [], folders = [];
  function File(name, mime, parent) { this.id = 'F' + (++fid); this.name = name; this.mime = mime; this.parent = parent; this.trashed = false; this.created = new Date(clock += 60000); }
  File.prototype.getName = function () { return this.name; }; File.prototype.setName = function (n) { this.name = n; return this; }; File.prototype.getMimeType = function () { return this.mime; };
  File.prototype.getDateCreated = function () { return this.created; }; File.prototype.setTrashed = function (t) { this.trashed = t; };
  File.prototype.getUrl = function () { return 'https://docs/' + this.id; };
  File.prototype.getId = function () { return this.id; };
  File.prototype.getLastUpdated = function () { return this.updated || this.created; };
  // вміст файла (виписки): рядок або Buffer; кодування, як у Таблицях: UTF-8 (з \uFFFD для чужих байтів) чи windows-1251
  File.prototype.getBlob = function () { const c = this.content; return { getDataAsString: cs => {
    const buf = Buffer.isBuffer(c) ? c : Buffer.from(String(c || ''), 'utf8');
    return !cs || /utf-?8/i.test(cs) ? new TextDecoder('utf-8').decode(buf) : new TextDecoder(String(cs).toLowerCase()).decode(buf); } }; };
  File.prototype.getParents = function () { const p = this.parent; let u = false; return { hasNext: () => !u && !!p, next: () => { u = true; return p; } }; };
  File.prototype.makeCopy = function (name, folder) { const f = new File(name, this.mime, folder); files.push(f); return f; };
  File.prototype.moveTo = function (folder) { this.parent = folder; return this; };
  File.prototype.isTrashed = function () { return this.trashed; };
  function Folder(name, parent) { this.id = 'D' + (++fid); this.name = name; this.parent = parent; this.trashed = false; folders.push(this); }
  Folder.prototype.getId = function () { return this.id; }; Folder.prototype.getUrl = function () { return 'https://drive/' + this.id; };
  Folder.prototype.isTrashed = function () { return this.trashed; }; Folder.prototype.createFolder = function (n) { return new Folder(n, this); };
  Folder.prototype.getFoldersByName = function (n) { const l = folders.filter(f => f.parent === this && f.name === n); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; };
  Folder.prototype.getFiles = function () { const l = files.filter(f => f.parent === this && !f.trashed); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; };
  const root = new Folder('root', null), ssFile = new File('Test', 'application/vnd.google-apps.spreadsheet', root);
  const addFile = (folder, name, mime, content) => { const f = new File(name, mime, folder); f.content = content; files.push(f); return f; };
  const others = {};
  const addSheetFile = (folder, name, rows, sheetNames) => {
    const f = addFile(folder, name, 'application/vnd.google-apps.spreadsheet', '');
    const names = sheetNames || ['Аркуш1'];
    others[f.id] = { getSheetByName: n => (names.indexOf(n) >= 0 ? {} : null),
      getSheets: () => [{ getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }), setValue: () => { throw new Error('write!'); } }] };
    return f;
  };
  // ---------- Documents (знімок для радника) ----------
  const docs = {};
  function Doc(name) {
    const f = new File(name, 'application/vnd.google-apps.document', root); files.push(f); this.file = f; this.paras = [{ t: '', h: 'NORMAL' }]; this.saved = 0; docs[f.id] = this;
  }
  Doc.prototype.getId = function () { return this.file.id; };
  Doc.prototype.getUrl = function () { return 'https://docs.google.com/document/d/' + this.file.id + '/edit'; };
  Doc.prototype.getName = function () { return this.file.name; };
  Doc.prototype.saveAndClose = function () { this.saved++; };
  Doc.prototype.getText = function () { return this.paras.map(p => p.t).join('\n'); };
  Doc.prototype.getBody = function () {
    const d = this, para = p => ({ setText: t => (p.t = String(t), para(p)), setHeading: h => (p.h = h, para(p)), getText: () => p.t });
    return { clear: () => { d.paras = [{ t: '', h: 'NORMAL' }]; }, getParagraphs: () => d.paras.map(para),
      appendParagraph: t => { const p = { t: String(t), h: 'NORMAL' }; d.paras.push(p); return para(p); }, getText: () => d.getText() };
  };
  const DocumentApp = {
    create: n => { if (opts.docFail) throw new Error(opts.docFail); calls.docs = (calls.docs || 0) + 1; return new Doc(n); },
    openById: id => { if (!docs[id] || docs[id].file.trashed) throw new Error('no document ' + id); return docs[id]; },
    ParagraphHeading: { NORMAL: 'NORMAL', HEADING1: 'HEADING1', HEADING2: 'HEADING2', HEADING3: 'HEADING3' },
  };
  const DriveApp = { getFileById: id => (id === 'SS1' ? ssFile : files.find(f => f.id === id)),
    getFolderById: id => { const f = folders.find(x => x.id === id); if (!f) throw new Error('no folder'); return f; }, getRootFolder: () => root };
  // ---------- Triggers / props / lock ----------
  const triggers = []; let tid = 0;
  const ScriptApp = { getProjectTriggers: () => triggers.slice(), deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
    newTrigger: h => { const spec = {}; const b = { timeBased: () => b, atHour: x => (spec.hour = x, b), nearMinute: x => (spec.minute = x, b), everyDays: x => (spec.days = x, b), after: x => (spec.after = x, b),
      everyMinutes: x => (spec.min = x, b), inTimezone: z => (spec.tz = z, b), forSpreadsheet: () => b, onEdit: () => (spec.onEdit = true, b), create: () => { const id = 'T' + (++tid); const t = { getHandlerFunction: () => h, getUniqueId: () => id, spec }; triggers.push(t); return t; } }; return b; } };
  const mkProps = () => { const st = {}; return { st, getProperty: k => (k in st ? st[k] : null), setProperty: (k, v) => { st[k] = String(v); },
    setProperties: o => Object.keys(o).forEach(k => (st[k] = String(o[k]))), getProperties: () => Object.assign({}, st), deleteProperty: k => { delete st[k]; } }; };
  const docProps = mkProps(), userProps = mkProps();
  let locked = false, docLocked = false;
  // ---------- Utilities (real time-zone handling) ----------
  function parts(d, zone) {
    const f = new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' });
    const o = {}; f.formatToParts(d).forEach(p => (o[p.type] = p.value)); return o;
  }
  const Utilities = {
    formatDate: (d, zone, fmt) => { const p = parts(d, zone); return fmt.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute).replace('ss', p.second); },
    parseDate: (s, zone, fmt) => {
      const m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}):(\d{2}))?$/);
      const guess = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
      const p = parts(new Date(guess), zone);
      const asUtc = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second);
      return new Date(guess - (asUtc - guess));
    },
    sleep: ms => { calls.sleep += ms; },
    getUuid: () => require('crypto').randomUUID(),
    DigestAlgorithm: { MD5: 'MD5' }, Charset: { UTF_8: 'UTF_8' },
    computeDigest: (alg, str) => Array.from(require('crypto').createHash('md5').update(String(str), 'utf8').digest()).map(b => (b > 127 ? b - 256 : b)),
    parseCsv: (text, delim) => {                                    // як Utilities.parseCsv: лапки, подвоєні лапки, переноси рядків у лапках
      const d = delim || ',', out = []; let row = [], cell = '', q = false;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; continue; }
        if (ch === '"') q = true; else if (ch === d) { row.push(cell); cell = ''; }
        else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); out.push(row); row = []; cell = ''; }
        else cell += ch;
      }
      if (cell !== '' || row.length) { row.push(cell); out.push(row); }
      return out;
    },
  };
  // ---------- Calendar (фаза E): календарі, події на цілий день, теги, нагадування ----------
  const scriptTz = opts.scriptTz || tz, cals = [];
  let cid = 0, eid = 0;
  const dayKey = d => Utilities.formatDate(d, scriptTz, 'yyyy-MM-dd');
  function CalEvent(cal, title, date, o) {
    this.cal = cal; this.id = 'E' + (++eid); this.title = title; this.day = dayKey(date); this.desc = (o && o.description) || '';
    this.tags = {}; this.rem = (cal.defaultRem || []).slice(); this.gone = false;
  }
  const E = CalEvent.prototype;
  E.getTitle = function () { return this.title; }; E.setTitle = function (t) { this.title = t; calls.cal.push(['title', this.id, t]); return this; };
  E.getDescription = function () { return this.desc; }; E.setDescription = function (d) { this.desc = d; return this; };
  E.getAllDayStartDate = function () { return Utilities.parseDate(this.day, scriptTz, 'yyyy-MM-dd'); };
  E.isAllDayEvent = function () { return true; };
  E.setAllDayDate = function (d) { this.day = dayKey(d); calls.cal.push(['date', this.id, this.day]); return this; };
  E.getTag = function (k) { return k in this.tags ? this.tags[k] : null; }; E.setTag = function (k, v) { this.tags[k] = String(v); return this; };
  E.addPopupReminder = function (m) { this.rem.push(m); return this; }; E.removeAllReminders = function () { this.rem = []; return this; };
  E.getPopupReminders = function () { return this.rem.slice(); };
  E.deleteEvent = function () { this.gone = true; calls.cal.push(['delete', this.id]); };
  E.getId = function () { return this.id; };
  function Cal(name, o) { this.id = 'C' + (++cid) + '@group.calendar.google.com'; this.name = name; this.summary = (o && o.summary) || ''; this.events = []; this.deleted = false; this.defaultRem = [30]; cals.push(this); }
  Cal.prototype.getId = function () { return this.id; }; Cal.prototype.getName = function () { return this.name; };
  Cal.prototype.deleteCalendar = function () { this.deleted = true; };
  Cal.prototype.getEvents = function (a, b) { const k1 = dayKey(a), k2 = dayKey(b); calls.cal.push(['getEvents', k1, k2]);
    return this.events.filter(e => !e.gone && e.day >= k1 && e.day <= k2); };
  Cal.prototype.createAllDayEvent = function (t, d, o) { if (opts.calFail) throw new Error(opts.calFail); const e = new CalEvent(this, t, d, o); this.events.push(e); calls.cal.push(['create', e.id, t, e.day]); return e; };
  const CalendarApp = {
    getCalendarById: id => { if (opts.calFail) throw new Error(opts.calFail); return cals.find(c => c.id === id && !c.deleted) || null; },
    getCalendarsByName: n => cals.filter(c => c.name === n && !c.deleted),
    createCalendar: (n, o) => { if (opts.calFail) throw new Error(opts.calFail); return new Cal(n, o); },
  };
  calls.cal = [];
  // ---------- UI ----------
  const ui = {
    Button: { YES: 'YES', NO: 'NO', OK: 'OK', CANCEL: 'CANCEL', CLOSE: 'CLOSE' }, ButtonSet: { YES_NO: 'YN', OK_CANCEL: 'OC', YES_NO_CANCEL: 'YNC' },
    alert: (t, m) => { calls.alerts.push([t, m]); const a = opts.alertAnswer; return typeof a === 'function' ? a(t, m) : (a || 'YES'); },
    // відповіді: рядок (на всі вікна однаково) або масив (по черзі; закінчились: порожньо); null у масиві: «Скасувати»
    prompt: (t, m) => { calls.prompts.push([t, m]);
      const a = Array.isArray(opts.promptText) ? (opts.promptText.length ? opts.promptText.shift() : '') : (opts.promptText || '');
      return { getSelectedButton: () => (a === null ? 'CANCEL' : opts.promptButton || 'OK'), getResponseText: () => (a === null ? '' : a) }; },
    showModalDialog: (h, t) => calls.dialogs.push({ t, html: h.html }),
    showSidebar: h => { calls.sidebars = (calls.sidebars || []).concat([{ t: h.title, html: h.html }]); },
    createMenu: () => { const m = { items: [], addItem: (a, b) => (m.items.push(b), m), addSeparator: () => m, addSubMenu: s => (m.items.push(...s.items), m), addToUi: () => (calls.menu = m.items) }; return m; },
  };
  // ---------- UrlFetch ----------
  const UrlFetchApp = { fetch: (url, o) => {
    calls.fetch.push(url);
    const r = (opts.mono || (() => ({ code: 500, body: 'no handler' })))(url, o);
    return { getResponseCode: () => r.code, getContentText: () => (typeof r.body === 'string' ? r.body : JSON.stringify(r.body)) };
  } };
  const ctx = {
    Charts: { ChartType: { LINE: 'LINE', COLUMN: 'COLUMN', AREA: 'AREA', BAR: 'BAR', PIE: 'PIE', COMBO: 'COMBO' } },
    SpreadsheetApp: { getActive: () => ss, flush: () => {}, openById: id => { if (!others[id]) throw new Error('no spreadsheet ' + id); calls.opened = (calls.opened || []).concat([id]); return others[id]; }, getUi: () => { if (opts.noUi) throw new Error('no ui'); return ui; },
      ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' }, newDataValidation: () => new DVB(), newConditionalFormatRule: () => new CFB(),
      DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST', VALUE_IN_RANGE: 'VALUE_IN_RANGE' },
      CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT', PASTE_DATA_VALIDATION: 'PASTE_DATA_VALIDATION' } },
    HtmlService: { createHtmlOutput: html => ({ html, setWidth() { return this; }, setHeight() { return this; }, setTitle(t) { this.title = t; return this; } }) },
    DriveApp, DocumentApp, ScriptApp, UrlFetchApp, Utilities, CalendarApp, Session: { getScriptTimeZone: () => scriptTz },
    PropertiesService: { getDocumentProperties: () => docProps, getUserProperties: () => userProps },
    LockService: { getScriptLock: () => ({ tryLock: () => { if (opts.lockBusy || locked) return false; locked = true; return true; }, releaseLock: () => { locked = false; } }),
      getDocumentLock: () => ({ tryLock: () => { if (opts.docLockBusy || docLocked) return false; docLocked = true; return true; }, releaseLock: () => { docLocked = false; } }) },
    Logger: { log: m => calls.logs.push(m) },
    console,
  };
  vm.createContext(ctx);
  if (opts.code) vm.runInContext(opts.code, ctx, { filename: 'Code.gs' }); else require('./source').load(ctx);
  return { ctx, ss, sheets, Sheet: (n) => ss.insertSheet(n), calls, docProps, userProps, triggers, files, isLocked: () => locked, Utilities, tz, colNum, colStr, cals,
    addFile, addSheetFile, root, docs,
    // інша таблиця (наприклад, стара копія шаблону з іншого makeEnv) для SpreadsheetApp.openById
    addSpreadsheet: (id, other) => { others[id] = other; } };
}
module.exports = { makeEnv, colNum, colStr, parseA1 };
