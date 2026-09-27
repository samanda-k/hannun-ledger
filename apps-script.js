/**
 * 한눈 가계부 · 구글 시트 서버 코드
 *
 * 이 코드 전체를 Apps Script 편집기(Code.gs)에 붙여넣고 저장하세요.
 *  1) 위쪽 함수 목록에서 setup 을 고르고 [실행] → 권한 허용
 *  2) [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → [배포]
 *  3) 웹 앱 URL과 실행 로그의 비밀코드를 가계부 앱에 입력
 *
 * 코드를 나중에 바꿨다면 [배포] → [배포 관리] → 연필 → 버전: 새 버전 → [배포] 를 해야 반영돼요.
 */

const ENTRY_SHEET = '내역';
const VISIT_SHEET = '방문';
const HIST_SHEET = '과거기록';
const SUMMARY_SHEET = '연간요약';
const ENTRY_HEAD = ['ID', '날짜', '월', '구분', '개인/사업', '분류', '결제수단', '공급가액', '부가세', '합계', '메모', '기록시각'];
const VISIT_HEAD = ['날짜', '월', '방문자수', '수정시각'];
const HIST_HEAD = ['연도', '월', '수입', '방문'];
const TYPE_TO_KO = { in: '수입', out: '지출', move: '저축·빚' };
const KO_TO_TYPE = { '수입': 'in', '지출': 'out', '저축·빚': 'move', '매출': 'in' };
const TOKEN_KEY = 'LEDGER_TOKEN';

/* ───────── 메뉴 ───────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('한눈 가계부')
    .addItem('비밀코드 보기', 'showToken')
    .addItem('비밀코드 새로 만들기', 'resetToken')
    .addSeparator()
    .addItem('연간요약 표 다시 만들기', 'buildSummary')
    .addToUi();
}

function showToken() {
  const t = getToken_();
  SpreadsheetApp.getUi().alert(t ? '가계부 앱에 입력할 비밀코드\n\n' + t : '아직 비밀코드가 없어요. Apps Script에서 setup 을 먼저 실행해 주세요.');
}

function resetToken() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('비밀코드를 새로 만들면 폰과 PC에서 다시 연결해야 해요. 계속할까요?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  ui.alert('새 비밀코드\n\n' + newToken_());
}

/* ───────── 처음 설정 ───────── */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Seoul');

  const e = ensureSheet_(ENTRY_SHEET, ENTRY_HEAD, [110, 95, 75, 70, 75, 90, 75, 95, 85, 95, 220, 140]);
  e.getRange('A:G').setNumberFormat('@');
  e.getRange('K:K').setNumberFormat('@');
  e.getRange('H:J').setNumberFormat('#,##0');
  e.getRange('L:L').setNumberFormat('yyyy-mm-dd hh:mm');

  const v = ensureSheet_(VISIT_SHEET, VISIT_HEAD, [95, 75, 80, 140]);
  v.getRange('A:B').setNumberFormat('@');
  v.getRange('D:D').setNumberFormat('yyyy-mm-dd hh:mm');

  const h = ensureSheet_(HIST_SHEET, HIST_HEAD, [70, 50, 110, 70]);
  h.getRange('C:C').setNumberFormat('#,##0');

  buildSummary();
  removeBlankDefaultSheet_();

  const token = getToken_() || newToken_();
  console.log('설정이 끝났어요. 가계부 앱에 입력할 비밀코드: ' + token);
  console.log('다음 단계: 오른쪽 위 [배포] → [새 배포] → 웹 앱');
}

function ensureSheet_(name, head, widths) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const first = sh.getRange(1, 1, 1, head.length);
  if (!sh.getRange(1, 1).getValue()) first.setValues([head]);
  first.setFontWeight('bold').setBackground('#1C5A4C').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  return sh;
}

function removeBlankDefaultSheet_() {
  const ss = SpreadsheetApp.getActive();
  ['시트1', 'Sheet1'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}

/* ───────── 비밀코드 ───────── */

function getToken_() {
  return PropertiesService.getScriptProperties().getProperty(TOKEN_KEY);
}

function newToken_() {
  const t = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty(TOKEN_KEY, t);
  return t;
}

/* ───────── 웹 앱 ───────── */

function doGet() {
  return json_({ ok: true, message: '한눈 가계부 서버가 켜져 있어요' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  const token = getToken_();
  if (!token || req.token !== token) return json_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'busy' });
  try {
    if (req.action === 'list') {
      return json_({
        ok: true,
        entries: readEntries_(),
        visits: readVisits_(),
        hist: readHist_(),
        sheetUrl: SpreadsheetApp.getActive().getUrl()
      });
    }
    if (req.action === 'batch') {
      applyOps_(Array.isArray(req.ops) ? req.ops : []);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    const msg = String(err && err.message || err);
    return json_({ ok: false, error: msg.indexOf('setup_needed') >= 0 ? 'setup_needed' : 'server', message: msg });
  } finally {
    lock.releaseLock();
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 읽기 ───────── */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('setup_needed');
  return sh;
}

function readEntries_() {
  const sh = sheet_(ENTRY_SHEET);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const range = sh.getRange(2, 1, n - 1, ENTRY_HEAD.length);
  const rows = range.getValues();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const out = [];
  let dirty = false;
  rows.forEach(r => {
    const d = normDate_(r[1], tz);
    const t = KO_TO_TYPE[String(r[3]).trim()];
    const amt = Math.abs(num_(r[9]));
    if (!d || !t || !amt) return;
    // 시트에 직접 적은 줄도 앱에서 보이도록 ID와 월을 채워 넣어요
    if (!r[0]) { r[0] = newId_(); dirty = true; }
    if (r[1] !== d) { r[1] = d; dirty = true; }
    if (r[2] !== d.slice(0, 7)) { r[2] = d.slice(0, 7); dirty = true; }
    const biz = String(r[4]).trim() === '사업' || String(r[3]).trim() === '매출';
    out.push({
      id: String(r[0]), d: d, t: t, biz: biz,
      cat: String(r[5]), pay: String(r[6]),
      sup: num_(r[7]), vat: num_(r[8]), amt: amt,
      memo: String(r[10]),
      ts: r[11] instanceof Date ? r[11].getTime() : num_(r[11])
    });
  });
  if (dirty) range.setValues(rows);
  return out;
}

function readVisits_() {
  const sh = sheet_(VISIT_SHEET);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const range = sh.getRange(2, 1, n - 1, VISIT_HEAD.length);
  const rows = range.getValues();
  const out = [];
  let dirty = false;
  rows.forEach(r => {
    const d = normDate_(r[0], tz);
    const c = num_(r[2]);
    if (!d || !c) return;
    if (r[0] !== d) { r[0] = d; dirty = true; }
    if (r[1] !== d.slice(0, 7)) { r[1] = d.slice(0, 7); dirty = true; }
    out.push({ d: d, n: c });
  });
  if (dirty) range.setValues(rows);
  return out;
}

function readHist_() {
  const sh = sheet_(HIST_SHEET);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HIST_HEAD.length).getValues()
    .map(r => ({ y: num_(r[0]), m: num_(r[1]), inc: num_(r[2]), vis: num_(r[3]) }))
    .filter(h => h.y > 1900 && h.m >= 1 && h.m <= 12 && (h.inc || h.vis));
}

/* ───────── 쓰기 ───────── */

function applyOps_(ops) {
  // 같은 대상에 여러 번 바뀐 경우 마지막 것만 반영해요
  const entries = {}, visits = {}, hist = {};
  ops.forEach(op => {
    if (!op || typeof op !== 'object') return;
    if (op.op === 'upsert' && op.entry && op.entry.id) entries[String(op.entry.id)] = { entry: op.entry };
    else if (op.op === 'delete' && op.id) entries[String(op.id)] = { del: true };
    else if (op.op === 'visits' && op.d) visits[normDate_(op.d, 'Asia/Seoul')] = Math.max(0, Math.round(num_(op.n)));
    else if (op.op === 'hist' && op.y && op.m) hist[num_(op.y) + '-' + num_(op.m)] ={ y: num_(op.y), m: num_(op.m), inc: num_(op.inc), vis: num_(op.vis) };
  });
  applyEntries_(entries);
  applyVisits_(visits);
  applyHist_(hist);
}

function applyEntries_(fin) {
  const ids = Object.keys(fin);
  if (!ids.length) return;
  const sh = sheet_(ENTRY_SHEET);
  const rowOf = indexRows_(sh, r => String(r[0]));
  const appends = [], dels = [];
  ids.forEach(id => {
    const f = fin[id], r = rowOf[id];
    if (f.del) { if (r) dels.push(r); return; }
    const row = toRow_(f.entry);
    if (!row) return;
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  });
  writeRows_(sh, appends, dels, ENTRY_HEAD.length);
  sortRows_(sh, [{ column: 2, ascending: true }, { column: 12, ascending: true }]);
}

function applyVisits_(fin) {
  const days = Object.keys(fin).filter(Boolean);
  if (!days.length) return;
  const sh = sheet_(VISIT_SHEET);
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const rowOf = indexRows_(sh, r => normDate_(r[0], tz));
  const appends = [], dels = [];
  days.forEach(d => {
    const n = fin[d], r = rowOf[d];
    if (!n) { if (r) dels.push(r); return; }
    const row = [d, d.slice(0, 7), n, new Date()];
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  });
  writeRows_(sh, appends, dels, VISIT_HEAD.length);
  sortRows_(sh, [{ column: 1, ascending: true }]);
}

function applyHist_(fin) {
  const keys = Object.keys(fin);
  if (!keys.length) return;
  const sh = sheet_(HIST_SHEET);
  const rowOf = indexRows_(sh, r => num_(r[0]) + '-' + num_(r[1]));
  const appends = [], dels = [];
  keys.forEach(k => {
    const h = fin[k], r = rowOf[k];
    if (!h.inc && !h.vis) { if (r) dels.push(r); return; }
    const row = [h.y, h.m, h.inc || '', h.vis || ''];
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  });
  writeRows_(sh, appends, dels, HIST_HEAD.length);
  sortRows_(sh, [{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}

function indexRows_(sh, keyOf) {
  const n = sh.getLastRow();
  const map = {};
  if (n < 2) return map;
  sh.getRange(2, 1, n - 1, 2).getValues().forEach((r, i) => {
    const k = keyOf(r);
    if (k) map[k] = i + 2;
  });
  return map;
}

function writeRows_(sh, appends, dels, width) {
  // 빈 줄을 넉넉히 두어 행 삭제·추가가 시트 크기에 막히지 않게 해요
  const need = sh.getLastRow() + appends.length + 20;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 100);
  dels.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
}

function sortRows_(sh, spec) {
  const n = sh.getLastRow();
  if (n > 2) sh.getRange(2, 1, n - 1, sh.getLastColumn()).sort(spec);
}

function toRow_(e) {
  const d = normDate_(e.d, 'Asia/Seoul');
  const t = TYPE_TO_KO[e.t];
  const amt = Math.abs(Math.round(num_(e.amt)));
  if (!d || !t || !amt) return null;
  const sup = Math.round(num_(e.sup)), vat = Math.round(num_(e.vat));
  return [
    String(e.id).slice(0, 40), d, d.slice(0, 7), t, e.biz ? '사업' : '개인',
    clean_(e.cat, 30), clean_(e.pay, 20),
    sup || '', vat || '', amt,
    clean_(e.memo, 200), new Date(num_(e.ts) || Date.now())
  ];
}

/* ───────── 도우미 ───────── */

function clean_(s, max) {
  s = String(s == null ? '' : s).slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = Number(String(v == null ? '' : v).replace(/[^\d.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function normDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(v == null ? '' : v).trim().match(/^(\d{2,4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/);
  if (!m) return '';
  const y = m[1].length === 2 ? '20' + m[1] : m[1];
  return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

/* ───────── 연간요약 표 (엑셀에서 쓰던 모양 그대로) ───────── */

function buildSummary() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET, 0);
  sh.clear();

  sh.getRange('A1:C1').setValues([['연도', new Date().getFullYear(), '← 연도를 바꾸면 표 전체가 바뀌어요']]);
  sh.getRange('A1').setFontWeight('bold');
  sh.getRange('B1').setNumberFormat('0').setFontWeight('bold').setBackground('#FFF2CC');
  sh.getRange('C1').setFontColor('#888888');

  const cols = 'BCDEFGHIJKLM'.split('');
  const head = ['구분'].concat(cols.map((c, i) => (i + 1) + '월'), ['합계']);
  sh.getRange(3, 1, 1, 14).setValues([head]).setFontWeight('bold').setBackground('#D9D9D9').setHorizontalAlignment('center');

  const E = "'" + ENTRY_SHEET + "'!", V = "'" + VISIT_SHEET + "'!", H = "'" + HIST_SHEET + "'!";
  const mm = m => (m < 10 ? '0' : '') + m;
  const key = (ye, m) => '(' + ye + ')&"-' + mm(m) + '"';
  const sp = parts => 'SUMPRODUCT(' + parts.join(',') + ')';
  const detail = (ye, m, t, biz, cat) => {
    const p = ['--(' + E + '$C$2:$C=' + key(ye, m) + ')', '--(' + E + '$D$2:$D="' + t + '")'];
    if (biz) p.push('--(' + E + '$E$2:$E="' + biz + '")');
    if (cat) p.push('--(' + E + '$F$2:$F="' + cat + '")');
    p.push(E + '$J$2:$J');
    return sp(p);
  };
  const histv = (ye, m, col) => sp(['--(' + H + '$A$2:$A=' + ye + ')', '--(' + H + '$B$2:$B=' + m + ')', H + '$' + col + '$2:$' + col]);
  const sales = (ye, m) => 'LET(dv,' + detail(ye, m, '수입', '사업') + ',IF(dv<>0,dv,' + histv(ye, m, 'C') + '))';
  const visits = (ye, m) => 'LET(dv,' + sp(['--(' + V + '$B$2:$B=' + key(ye, m) + ')', V + '$C$2:$C']) + ',IF(dv<>0,dv,' + histv(ye, m, 'D') + '))';

  // 줄 정의: base = 달마다 계산(합계 열은 SUM), derived = 같은 열의 다른 줄로 계산
  const Y = '$B$1';
  const rows = [
    { id: 'y3', label: '=(' + Y + '-3)&"년 수입"', base: m => sales(Y + '-3', m) },
    { id: 'y2', label: '=(' + Y + '-2)&"년 수입"', base: m => sales(Y + '-2', m) },
    { id: 'y1', label: '=(' + Y + '-1)&"년 수입"', base: m => sales(Y + '-1', m) },
    { id: 'v1', label: '=(' + Y + '-1)&"년 방문"', base: m => visits(Y + '-1', m) },
    { id: 'y0', label: '=' + Y + '&"년 수입"', base: m => sales(Y, m), bg: '#FFF2CC', bold: true },
    { id: 'v0', label: '=' + Y + '&"년 방문"', base: m => visits(Y, m), bg: '#FFF2CC', bold: true },
    { id: 'etc', label: '기타 수입', base: m => detail(Y, m, '수입', '개인'), bg: '#FFF2CC' },
    { id: 'yoy', label: '작년대비', der: c => 'IF(' + c + '{y1}=0,"",(' + c + '{y0}-' + c + '{y1})/' + c + '{y1})', pct: true, bg: '#FCE4D6' },
    { id: 'fix', label: '고정비', base: m => detail(Y, m, '지출', '사업', '고정비'), bg: '#DDEBF7' },
    { id: 'sup', label: '소모품', base: m => detail(Y, m, '지출', '사업', '소모품'), bg: '#FFF2CC' },
    { id: 'biz', label: '사업자', der: c => c + '{bot}-' + c + '{fix}-' + c + '{sup}', bg: '#E2EFDA' },
    { id: 'bot', label: '지출 계', base: m => detail(Y, m, '지출', '사업'), bg: '#FCE4D6', bold: true },
    { id: 'sal', label: '월 급여', der: c => c + '{y0}-' + c + '{bot}', bg: '#F8CBAD', bold: true },
    { id: 'mar', label: '수익률', der: c => 'IF(' + c + '{y0}=0,"",' + c + '{sal}/' + c + '{y0})', pct: true, bg: '#FFFF99', bold: true },
    { id: 'per', label: '개인지출', base: m => detail(Y, m, '지출', '개인'), bg: '#E2EFDA' },
    { id: 'bor', label: '빌린 돈', base: m => detail(Y, m, '저축·빚', '', '빌린 돈'), bg: '#FCE4D6' },
    { id: 'new', label: '새출발기금', base: m => detail(Y, m, '저축·빚', '', '새출발기금'), bg: '#DDEBF7' },
    { id: 'sav', label: '모으기', base: m => detail(Y, m, '저축·빚', '', '모으기'), bg: '#E4DFEC' },
    { id: 'left', label: '남는 돈', der: c => c + '{sal}+' + c + '{etc}+' + c + '{bor}-' + c + '{per}-' + c + '{new}-' + c + '{sav}', bg: '#FFF2CC', bold: true }
  ];
  const START = 4;
  const rowNo = {};
  rows.forEach((r, i) => { rowNo[r.id] = START + i; });
  const fill = (tpl, c) => tpl.replace(/\{(\w+)\}/g, (_, id) => String(rowNo[id]));

  const labels = [], formulas = [];
  rows.forEach((r, i) => {
    const rn = START + i;
    labels.push([r.label]);
    const line = cols.map((c, mi) => '=' + (r.base ? r.base(mi + 1) : fill(r.der(c), c)));
    line.push('=' + (r.base ? 'SUM(B' + rn + ':M' + rn + ')' : fill(r.der('N'), 'N')));
    formulas.push(line);
  });
  sh.getRange(START, 1, rows.length, 1).setValues(labels);
  sh.getRange(START, 2, rows.length, 13).setFormulas(formulas);

  rows.forEach((r, i) => {
    const rg = sh.getRange(START + i, 1, 1, 14);
    rg.setNumberFormat(r.pct ? '0%' : '#,##0;-#,##0;"-"');
    if (r.bg) rg.setBackground(r.bg);
    if (r.bold) rg.setFontWeight('bold');
  });
  sh.getRange(START, 1, rows.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(START, 2, rows.length, 13).setHorizontalAlignment('right');
  sh.getRange(3, 1, rows.length + 1, 14).setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 110);
  for (let c = 2; c <= 14; c++) sh.setColumnWidth(c, 95);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}
