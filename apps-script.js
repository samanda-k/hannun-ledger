/**
 * 한눈 가계부 · 구글 시트 서버 코드 (버전 6)
 *
 * 이 코드 전체를 Apps Script 편집기(Code.gs)에 붙여넣고 저장하세요.
 *  1) 함수 목록에서 setup 을 고르고 [실행] → 권한 허용
 *     (또는 구글 시트를 새로고침한 뒤 메뉴 [한눈 가계부] → [처음 설정])
 *  2) [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → [배포]
 *  3) 웹 앱 URL과 비밀코드를 가계부 앱에 입력
 *
 * 코드를 나중에 바꿨다면 [배포] → [배포 관리] → 연필 → 버전: 새 버전 → [배포] 를 해야 반영돼요.
 */

var ENTRY_SHEET = '내역';
var VISIT_SHEET = '방문';
var HIST_SHEET = '과거기록';
var SUMMARY_SHEET = '연간요약';
var BAL_SHEET = '잔고';
var BAL_HEAD = ['월', '통장 잔고', '현금 잔고', '수정시각'];
var SERVER_VER = 6;
var ENTRY_HEAD = ['ID', '날짜', '월', '구분', '개인/사업', '분류', '결제수단', '공급가액', '부가세', '합계', '메모', '기록시각'];
var VISIT_HEAD = ['날짜', '월', '방문자수', '수정시각'];
var HIST_HEAD = ['연도', '월', '수입', '방문'];
var TYPE_TO_KO = { 'in': '수입', 'out': '지출', 'move': '저축·빚' };
var KO_TO_TYPE = { '수입': 'in', '지출': 'out', '저축·빚': 'move', '매출': 'in' };
var TOKEN_KEY = 'LEDGER_TOKEN';

/* ───────── 메뉴 ───────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('한눈 가계부')
    .addItem('처음 설정 (setup)', 'setup')
    .addItem('비밀코드 보기', 'showToken')
    .addItem('비밀코드 새로 만들기', 'resetToken')
    .addSeparator()
    .addItem('연간요약 표 다시 만들기', 'buildSummary')
    .addToUi();
}

function showToken() {
  var t = getToken_();
  SpreadsheetApp.getUi().alert(t ? '가계부 앱에 입력할 비밀코드\n\n' + t : '아직 비밀코드가 없어요. 메뉴 [한눈 가계부] → [처음 설정]을 먼저 눌러 주세요.');
}

function resetToken() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('비밀코드를 새로 만들면 폰과 PC에서 다시 연결해야 해요. 계속할까요?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  ui.alert('새 비밀코드\n\n' + newToken_());
}

/* ───────── 처음 설정 ───────── */

function setup() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('구글 시트에서 [확장 프로그램] → [Apps Script]로 연 편집기에 붙여넣어야 해요.');
  ss.setSpreadsheetTimeZone('Asia/Seoul');

  var e = ensureSheet_(ENTRY_SHEET, ENTRY_HEAD, [110, 95, 75, 70, 75, 90, 75, 95, 85, 95, 220, 140]);
  e.getRange('A:G').setNumberFormat('@');
  e.getRange('K:K').setNumberFormat('@');
  e.getRange('H:J').setNumberFormat('#,##0');
  e.getRange('L:L').setNumberFormat('yyyy-mm-dd hh:mm');

  var v = ensureSheet_(VISIT_SHEET, VISIT_HEAD, [95, 75, 80, 140]);
  v.getRange('A:B').setNumberFormat('@');
  v.getRange('D:D').setNumberFormat('yyyy-mm-dd hh:mm');

  var h = ensureSheet_(HIST_SHEET, HIST_HEAD, [70, 50, 110, 70]);
  h.getRange('C:C').setNumberFormat('#,##0');

  balSheet_();

  buildSummary();
  removeBlankDefaultSheet_();

  var token = getToken_() || newToken_();
  Logger.log('설정이 끝났어요. 가계부 앱에 입력할 비밀코드: ' + token);
  Logger.log('다음 단계: 오른쪽 위 [배포] → [새 배포] → 웹 앱');
  try {
    // 시트 메뉴에서 실행했을 때는 팝업으로 보여줘요 (편집기에서 실행하면 실행 로그에 나와요)
    SpreadsheetApp.getUi().alert('설정이 끝났어요.\n\n가계부 앱에 입력할 비밀코드\n' + token + '\n\n다음 단계: Apps Script 편집기에서 [배포] → [새 배포] → 웹 앱');
  } catch (err) {}
}

function ensureSheet_(name, head, widths) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var first = sh.getRange(1, 1, 1, head.length);
  if (!sh.getRange(1, 1).getValue()) first.setValues([head]);
  first.setFontWeight('bold').setBackground('#1C5A4C').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
  return sh;
}

function removeBlankDefaultSheet_() {
  var ss = SpreadsheetApp.getActive();
  var names = ['시트1', 'Sheet1'];
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  }
}

/* ───────── 비밀코드 ───────── */

function getToken_() {
  return PropertiesService.getScriptProperties().getProperty(TOKEN_KEY);
}

function newToken_() {
  var t = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty(TOKEN_KEY, t);
  return t;
}

/* ───────── 웹 앱 ───────── */

function doGet() {
  return json_({ ok: true, message: '한눈 가계부 서버가 켜져 있어요' });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  var token = getToken_();
  if (!token || req.token !== token) return json_({ ok: false, error: 'unauthorized' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'busy' });
  try {
    if (req.action === 'list') {
      return json_({
        ok: true,
        entries: readEntries_(),
        visits: readVisits_(),
        hist: readHist_(),
        bal: readBal_(),
        ver: SERVER_VER,
        sheetUrl: SpreadsheetApp.getActive().getUrl()
      });
    }
    if (req.action === 'batch') {
      applyOps_(req.ops instanceof Array ? req.ops : []);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    var msg = String(err && err.message || err);
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
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('setup_needed');
  return sh;
}

function readEntries_() {
  var sh = sheet_(ENTRY_SHEET);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var range = sh.getRange(2, 1, n - 1, ENTRY_HEAD.length);
  var rows = range.getValues();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var out = [];
  var dirty = false;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var d = normDate_(r[1], tz);
    var kind = String(r[3]).replace(/^\s+|\s+$/g, '');
    var t = KO_TO_TYPE[kind];
    var amt = Math.abs(num_(r[9]));
    if (!d || !t || !amt) continue;
    // 시트에 직접 적은 줄도 앱에서 보이도록 ID와 월을 채워 넣어요
    if (!r[0]) { r[0] = newId_(); dirty = true; }
    if (r[1] !== d) { r[1] = d; dirty = true; }
    if (r[2] !== d.slice(0, 7)) { r[2] = d.slice(0, 7); dirty = true; }
    var biz = String(r[4]).replace(/^\s+|\s+$/g, '') === '사업' || kind === '매출';
    out.push({
      id: String(r[0]), d: d, t: t, biz: biz,
      cat: String(r[5]), pay: String(r[6]),
      sup: num_(r[7]), vat: num_(r[8]), amt: amt,
      memo: String(r[10]),
      ts: r[11] instanceof Date ? r[11].getTime() : num_(r[11])
    });
  }
  if (dirty) range.setValues(rows);
  return out;
}

function readVisits_() {
  var sh = sheet_(VISIT_SHEET);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var range = sh.getRange(2, 1, n - 1, VISIT_HEAD.length);
  var rows = range.getValues();
  var out = [];
  var dirty = false;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var d = normDate_(r[0], tz);
    var c = num_(r[2]);
    if (!d || !c) continue;
    if (r[0] !== d) { r[0] = d; dirty = true; }
    if (r[1] !== d.slice(0, 7)) { r[1] = d.slice(0, 7); dirty = true; }
    out.push({ d: d, n: c });
  }
  if (dirty) range.setValues(rows);
  return out;
}

function readHist_() {
  var sh = sheet_(HIST_SHEET);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var rows = sh.getRange(2, 1, n - 1, HIST_HEAD.length).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var h = { y: num_(rows[i][0]), m: num_(rows[i][1]), inc: num_(rows[i][2]), vis: num_(rows[i][3]) };
    if (h.y > 1900 && h.m >= 1 && h.m <= 12 && (h.inc || h.vis)) out.push(h);
  }
  return out;
}

/* ───────── 쓰기 ───────── */

function applyOps_(ops) {
  // 같은 대상이 여러 번 바뀐 경우 마지막 것만 반영해요
  var entries = {}, visits = {}, hist = {}, bal = {};
  for (var i = 0; i < ops.length; i++) {
    var op = ops[i];
    if (!op || typeof op !== 'object') continue;
    if (op.op === 'upsert' && op.entry && op.entry.id) entries[String(op.entry.id)] = { entry: op.entry };
    else if (op.op === 'delete' && op.id) entries[String(op.id)] = { del: true };
    else if (op.op === 'visits' && op.d) visits[normDate_(op.d, 'Asia/Seoul')] = Math.max(0, Math.round(num_(op.n)));
    else if (op.op === 'bal' && /^\d{4}-\d{2}$/.test(String(op.ym))) bal[String(op.ym)] = { bank: num_(op.bank), cash: num_(op.cash), hasBank: op.bank !== null && op.bank !== '', hasCash: op.cash !== null && op.cash !== '' };
    else if (op.op === 'hist' && op.y && op.m) hist[num_(op.y) + '-' + num_(op.m)] = { y: num_(op.y), m: num_(op.m), inc: num_(op.inc), vis: num_(op.vis) };
  }
  applyEntries_(entries);
  applyVisits_(visits);
  applyHist_(hist);
  applyBal_(bal);
}

function applyEntries_(fin) {
  var ids = Object.keys(fin);
  if (!ids.length) return;
  var sh = sheet_(ENTRY_SHEET);
  var rowOf = indexRows_(sh, function (r) { return String(r[0]); });
  var appends = [], dels = [];
  for (var i = 0; i < ids.length; i++) {
    var f = fin[ids[i]], r = rowOf[ids[i]];
    if (f.del) { if (r) dels.push(r); continue; }
    var row = toRow_(f.entry);
    if (!row) continue;
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  }
  writeRows_(sh, appends, dels, ENTRY_HEAD.length);
  sortRows_(sh, [{ column: 2, ascending: true }, { column: 12, ascending: true }]);
}

function applyVisits_(fin) {
  var days = Object.keys(fin).filter(function (d) { return !!d; });
  if (!days.length) return;
  var sh = sheet_(VISIT_SHEET);
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var rowOf = indexRows_(sh, function (r) { return normDate_(r[0], tz); });
  var appends = [], dels = [];
  for (var i = 0; i < days.length; i++) {
    var d = days[i], n = fin[d], r = rowOf[d];
    if (!n) { if (r) dels.push(r); continue; }
    var row = [d, d.slice(0, 7), n, new Date()];
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  }
  writeRows_(sh, appends, dels, VISIT_HEAD.length);
  sortRows_(sh, [{ column: 1, ascending: true }]);
}

function applyHist_(fin) {
  var keys = Object.keys(fin);
  if (!keys.length) return;
  var sh = sheet_(HIST_SHEET);
  var rowOf = indexRows_(sh, function (r) { return num_(r[0]) + '-' + num_(r[1]); });
  var appends = [], dels = [];
  for (var i = 0; i < keys.length; i++) {
    var h = fin[keys[i]], r = rowOf[keys[i]];
    if (!h.inc && !h.vis) { if (r) dels.push(r); continue; }
    var row = [h.y, h.m, h.inc || '', h.vis || ''];
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  }
  writeRows_(sh, appends, dels, HIST_HEAD.length);
  sortRows_(sh, [{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}

function balSheet_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(BAL_SHEET);
  if (sh) return sh;
  sh = ensureSheet_(BAL_SHEET, BAL_HEAD, [80, 120, 120, 140]);
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('B:C').setNumberFormat('#,##0');
  sh.getRange('D:D').setNumberFormat('yyyy-mm-dd hh:mm');
  return sh;
}

function readBal_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(BAL_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, BAL_HEAD.length).getValues();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var ym = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM') : String(r[0]).replace(/^\s+|\s+$/g, '');
    if (!/^\d{4}-\d{2}$/.test(ym)) continue;
    out.push({ ym: ym, bank: r[1] === '' ? null : num_(r[1]), cash: r[2] === '' ? null : num_(r[2]),
      at: r[3] instanceof Date ? r[3].getTime() : 0 });
  }
  return out;
}

function applyBal_(fin) {
  var keys = Object.keys(fin);
  if (!keys.length) return;
  var sh = balSheet_();
  var rowOf = indexRows_(sh, function (r) { return r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Seoul', 'yyyy-MM') : String(r[0]); });
  var appends = [];
  for (var i = 0; i < keys.length; i++) {
    var b = fin[keys[i]], r = rowOf[keys[i]];
    var row = [keys[i], b.hasBank ? b.bank : '', b.hasCash ? b.cash : '', new Date()];
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else appends.push(row);
  }
  writeRows_(sh, appends, [], BAL_HEAD.length);
  sortRows_(sh, [{ column: 1, ascending: true }]);
}

function indexRows_(sh, keyOf) {
  var n = sh.getLastRow();
  var map = {};
  if (n < 2) return map;
  var rows = sh.getRange(2, 1, n - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    var k = keyOf(rows[i]);
    if (k) map[k] = i + 2;
  }
  return map;
}

function writeRows_(sh, appends, dels, width) {
  // 빈 줄을 넉넉히 두어 행 삭제·추가가 시트 크기에 막히지 않게 해요
  var need = sh.getLastRow() + appends.length + 20;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 100);
  dels.sort(function (a, b) { return b - a; });
  for (var i = 0; i < dels.length; i++) sh.deleteRow(dels[i]);
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
}

function sortRows_(sh, spec) {
  var n = sh.getLastRow();
  if (n > 2) sh.getRange(2, 1, n - 1, sh.getLastColumn()).sort(spec);
}

function toRow_(e) {
  var d = normDate_(e.d, 'Asia/Seoul');
  var t = TYPE_TO_KO[e.t];
  var amt = Math.abs(Math.round(num_(e.amt)));
  if (!d || !t || !amt) return null;
  var sup = Math.round(num_(e.sup)), vat = Math.round(num_(e.vat));
  return [
    String(e.id).slice(0, 40), d, d.slice(0, 7), t, e.biz ? '사업' : '개인',
    clean_(e.cat, 30), clean_(e.pay, 20),
    sup || '', vat || '', amt,
    clean_(e.memo, 200), new Date(num_(e.ts) || new Date().getTime())
  ];
}

/* ───────── 도우미 ───────── */

function clean_(s, max) {
  s = String(s === null || s === undefined ? '' : s).slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = Number(String(v === null || v === undefined ? '' : v).replace(/[^\d.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function pad2_(x) {
  return ('0' + x).slice(-2);
}

function normDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v === null || v === undefined ? '' : v).replace(/^\s+|\s+$/g, '');
  var m = s.match(/^(\d{2,4})\s*[-.\/년]\s*(\d{1,2})\s*[-.\/월]\s*(\d{1,2})/);
  if (!m) return '';
  var y = m[1].length === 2 ? '20' + m[1] : m[1];
  return y + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
}

/* ───────── 연간요약 표 (엑셀에서 쓰던 모양 그대로) ───────── */

function buildSummary() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET, 0);
  sh.clear();

  sh.getRange('A1:C1').setValues([['연도', new Date().getFullYear(), '← 연도를 바꾸면 표 전체가 바뀌어요']]);
  sh.getRange('A1').setFontWeight('bold');
  sh.getRange('B1').setNumberFormat('0').setFontWeight('bold').setBackground('#FFF2CC');
  sh.getRange('C1').setFontColor('#888888');

  var cols = 'BCDEFGHIJKLM'.split('');
  var head = ['구분'];
  for (var c = 0; c < 12; c++) head.push((c + 1) + '월');
  head.push('합계');
  sh.getRange(3, 1, 1, 14).setValues([head]).setFontWeight('bold').setBackground('#D9D9D9').setHorizontalAlignment('center');

  var E = "'" + ENTRY_SHEET + "'!", V = "'" + VISIT_SHEET + "'!", H = "'" + HIST_SHEET + "'!";
  function key(ye, m) { return '(' + ye + ')&"-' + pad2_(m) + '"'; }
  function sp(parts) { return 'SUMPRODUCT(' + parts.join(',') + ')'; }
  function detail(ye, m, t, biz, cat, notCat) {
    var p = ['--(' + E + '$C$2:$C=' + key(ye, m) + ')', '--(' + E + '$D$2:$D="' + t + '")'];
    if (biz) p.push('--(' + E + '$E$2:$E="' + biz + '")');
    if (cat) p.push('--(' + E + '$F$2:$F="' + cat + '")');
    if (notCat) p.push('--(' + E + '$F$2:$F<>"' + notCat + '")');
    p.push(E + '$J$2:$J');
    return sp(p);
  }
  function histv(ye, m, col) {
    return sp(['--(' + H + '$A$2:$A=' + ye + ')', '--(' + H + '$B$2:$B=' + m + ')', H + '$' + col + '$2:$' + col]);
  }
  function sales(ye, m) { return 'LET(dv,' + detail(ye, m, '수입', '사업') + ',IF(dv<>0,dv,' + histv(ye, m, 'C') + '))'; }
  function visits(ye, m) { return 'LET(dv,' + sp(['--(' + V + '$B$2:$B=' + key(ye, m) + ')', V + '$C$2:$C']) + ',IF(dv<>0,dv,' + histv(ye, m, 'D') + '))'; }

  // base = 달마다 계산(합계 열은 SUM), der = 같은 열의 다른 줄로 계산
  var Y = '$B$1';
  var rows = [
    { id: 'y3', label: '=(' + Y + '-3)&"년 수입"', base: function (m) { return sales(Y + '-3', m); } },
    { id: 'y2', label: '=(' + Y + '-2)&"년 수입"', base: function (m) { return sales(Y + '-2', m); } },
    { id: 'y1', label: '=(' + Y + '-1)&"년 수입"', base: function (m) { return sales(Y + '-1', m); } },
    { id: 'v1', label: '=(' + Y + '-1)&"년 방문"', base: function (m) { return visits(Y + '-1', m); } },
    { id: 'y0', label: '=' + Y + '&"년 수입"', base: function (m) { return sales(Y, m); }, bg: '#FFF2CC', bold: true },
    { id: 'v0', label: '=' + Y + '&"년 방문"', base: function (m) { return visits(Y, m); }, bg: '#FFF2CC', bold: true },
    { id: 'etc', label: '기타 수입', base: function (m) { return detail(Y, m, '수입', '개인', '', '대출·자금'); }, bg: '#FFF2CC' },
    { id: 'yoy', label: '작년대비', der: function (c) { return 'IF(' + c + '{y1}=0,"",(' + c + '{y0}-' + c + '{y1})/' + c + '{y1})'; }, pct: true, bg: '#FCE4D6' },
    { id: 'fix', label: '고정비', base: function (m) { return detail(Y, m, '지출', '사업', '고정비'); }, bg: '#DDEBF7' },
    { id: 'sup', label: '소모품', base: function (m) { return detail(Y, m, '지출', '사업', '소모품'); }, bg: '#FFF2CC' },
    { id: 'biz', label: '사업자', der: function (c) { return c + '{bot}-' + c + '{fix}-' + c + '{sup}'; }, bg: '#E2EFDA' },
    { id: 'bot', label: '지출 계', base: function (m) { return detail(Y, m, '지출', '사업', '', '창업·이전'); }, bg: '#FCE4D6', bold: true },
    { id: 'sal', label: '월 급여', der: function (c) { return c + '{y0}-' + c + '{bot}'; }, bg: '#F8CBAD', bold: true },
    { id: 'mar', label: '수익률', der: function (c) { return 'IF(' + c + '{y0}=0,"",' + c + '{sal}/' + c + '{y0})'; }, pct: true, bg: '#FFFF99', bold: true },
    { id: 'per', label: '개인지출', base: function (m) { return detail(Y, m, '지출', '개인'); }, bg: '#E2EFDA' },
    { id: 'bor', label: '빌린 돈 갚기', base: function (m) { return detail(Y, m, '저축·빚', '', '빌린 돈 갚기'); }, bg: '#FCE4D6' },
    { id: 'new', label: '새출발기금', base: function (m) { return detail(Y, m, '저축·빚', '', '새출발기금'); }, bg: '#DDEBF7' },
    { id: 'sav', label: '모으기', base: function (m) { return detail(Y, m, '저축·빚', '', '모으기'); }, bg: '#E4DFEC' },
    { id: 'yu', label: '노란우산공제', base: function (m) { return detail(Y, m, '저축·빚', '', '노란우산공제'); }, bg: '#FFF2CC' },
    { id: 'yuc', label: '노란우산 적립금', cum: function (m) { return sp(['--(' + E + '$C$2:$C<=' + key(Y, m) + ')', '--(' + E + '$D$2:$D="저축·빚")', '--(' + E + '$F$2:$F="노란우산공제")', E + '$J$2:$J']); }, bg: '#FFF2CC' },
    { id: 'left', label: '남는 돈', der: function (c) { return c + '{sal}+' + c + '{etc}-' + c + '{per}-' + c + '{bor}-' + c + '{new}-' + c + '{sav}-' + c + '{yu}'; }, bg: '#FFF2CC', bold: true },
    { id: 'stp', label: '창업·이전 비용', base: function (m) { return detail(Y, m, '지출', '사업', '창업·이전'); }, bg: '#EDEDED' },
    { id: 'fnd', label: '대출·자금', base: function (m) { return detail(Y, m, '수입', '개인', '대출·자금'); }, bg: '#EDEDED' }
  ];
  var START = 4;
  var rowNo = {};
  for (var i = 0; i < rows.length; i++) rowNo[rows[i].id] = START + i;
  function fill(tpl) { return tpl.replace(/\{(\w+)\}/g, function (_, id) { return String(rowNo[id]); }); }

  var labels = [], formulas = [];
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r], rn = START + r, line = [];
    labels.push([row.label]);
    for (var mi = 0; mi < 12; mi++) line.push('=' + (row.cum ? row.cum(mi + 1) : row.base ? row.base(mi + 1) : fill(row.der(cols[mi]))));
    line.push('=' + (row.cum ? 'M' + rn : row.base ? 'SUM(B' + rn + ':M' + rn + ')' : fill(row.der('N'))));
    formulas.push(line);
  }
  sh.getRange(START, 1, rows.length, 1).setValues(labels);
  sh.getRange(START, 2, rows.length, 13).setFormulas(formulas);

  for (var k = 0; k < rows.length; k++) {
    var rg = sh.getRange(START + k, 1, 1, 14);
    rg.setNumberFormat(rows[k].pct ? '0%' : '#,##0;-#,##0;"-"');
    if (rows[k].bg) rg.setBackground(rows[k].bg);
    if (rows[k].bold) rg.setFontWeight('bold');
  }
  sh.getRange(START, 1, rows.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(START, 2, rows.length, 13).setHorizontalAlignment('right');
  sh.getRange(3, 1, rows.length + 1, 14).setBorder(true, true, true, true, true, true, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 110);
  for (var cc = 2; cc <= 14; cc++) sh.setColumnWidth(cc, 95);
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);
}
