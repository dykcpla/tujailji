/* ===== 엑셀(투자일지 .xlsx) → 프로그램 데이터 변환 ===== */
var XIMP = (function () {
  function parse(X, buf, prev) {
    var wb = X.read(buf, { type: 'array', cellFormula: true, cellDates: false, cellStyles: false, cellHTML: false, cellText: false, cellNF: false, sheets: ['BOARD', '목표비중', '거래내역', '월말일현황', '배당', '자녀증여', 'PLAN', 'SIM', '2031-40추계'] });
    var need = ['BOARD', '거래내역', '월말일현황', '배당', '자녀증여', 'PLAN', 'SIM'];
    var miss = need.filter(function (n) { return !wb.Sheets[n]; });
    if (miss.length) throw new Error('시트가 없습니다: ' + miss.join(', '));
    function sh(n) { return wb.Sheets[n]; }
    function c(s, a) { return sh(s)[a]; }
    function v(s, a) { var x = c(s, a); return x ? x.v : null; }
    function num(s, a) { var x = v(s, a); return typeof x === 'number' ? x : (x == null || x === '' ? null : (isFinite(+x) ? +x : null)); }
    function f(s, a) { var x = c(s, a); return x && x.f ? String(x.f) : null; }
    function str(s, a) { var x = v(s, a); return x == null ? '' : String(x).trim(); }
    function iso(s, a) {
      var x = c(s, a); if (!x || x.v == null || x.v === '') return null;
      if (x.t === 'n') { var d = X.SSF.parse_date_code(x.v); if (!d || d.y < 1900) return null; return d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0'); }
      if (x.t === 'd') { var dd = new Date(x.v); return dd.getFullYear() + '-' + String(dd.getMonth() + 1).padStart(2, '0') + '-' + String(dd.getDate()).padStart(2, '0'); }
      var m = String(x.v).match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/); return m ? m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0') : null;
    }
    function A(col, row) { return col + row; }
    function lastRow(s) { var r = sh(s)['!ref']; return r ? X.utils.decode_range(r).e.r + 1 : 0; }
    function r6(x) { return typeof x === 'number' ? Math.round(x * 1e6) / 1e6 : x; }

    /* --- 설정: BOARD --- */
    var POS = [
      { col: 'D', fx: 'D9', price: 'E11', c: 'USD', g: 'SCHD', acct: '해외직투', dy: 3, ref: '$D$9' },
      { col: 'H', fx: 'H9', price: 'I11', c: 'USD', g: 'NSD100', acct: '해외직투', dy: 4, ref: '$H$9' },
      { col: 'L', fx: 'L9', price: 'M11', c: 'USD', g: 'S&P500', acct: '해외직투', dy: 5, ref: '$L$9' },
      { col: 'P', price: 'P11', c: 'KRW', g: 'NSD100', acct: 'ISA', dy: 6 },
      { col: 'S', price: 'S11', c: 'KRW', g: 'S&P500', acct: '연금저축', dy: 7 },
      { col: 'V', price: 'V11', c: 'KRW', g: 'GOLD', acct: '금현물계좌', dy: null },
      { col: 'Y', fx: 'Y9', price: 'Z11', c: 'USD', g: 'SPCX', acct: '해외직투', dy: null, ref: '$Y$9' }
    ];
    var prevT = {}; ((prev && prev.settings && prev.settings.tickers) || []).forEach(function (t) { prevT[t.t] = t; });
    var refMap = {};
    var tickers = POS.map(function (p) {
      var name = str('BOARD', p.col + '5'), old = prevT[name] || {};
      var t = { t: name, c: p.c, g: old.g || p.g, acct: old.acct || p.acct, price: num('BOARD', p.price) || 0, dy: 0, tax: p.c === 'USD' ? 0.15 : 0, inBuy: old.inBuy != null ? old.inBuy : p.g !== 'GOLD' };
      if (p.fx) { t.fx = num('BOARD', p.fx); refMap[p.ref] = name; }
      if (p.dy) { t.dy = num('BOARD', 'AL' + p.dy) || 0; t.tax = num('BOARD', 'AM' + p.dy) || 0; }
      else if (old.dy != null) { t.dy = old.dy; t.tax = old.tax; }
      return t;
    });
    // 프로그램에서 따로 추가한 종목은 유지
    ((prev && prev.settings && prev.settings.tickers) || []).forEach(function (t) { if (!tickers.some(function (x) { return x.t === t.t; })) tickers.push(t); });
    var pension = {}, isa = {};
    for (var r = 23; r <= 40; r++) {
      var lab = str('BOARD', 'AD' + r), y = (lab.match(/(\d{4})/) || [])[1]; if (!y) continue;
      pension[y] = num('BOARD', 'AF' + r) || 0; isa[y] = num('BOARD', 'AI' + r) || (+y >= 2026 ? 20000000 : 0);
    }
    var asOf = iso('BOARD', 'AY1');
    var ps = (prev && prev.settings) || {};
    var settings = {
      asOf: asOf, priceAt: asOf || (ps.priceAt || null), fxNow: num('BOARD', 'B8'), deposit: num('BOARD', 'L4') || 0, cashBal: num('BOARD', 'O4') || 0,
      bilUsd: num('BOARD', 'S4') || 0, bilFx: num('BOARD', 'R4') || 0, bilDy: num('BOARD', 'AL8') || 0, bilTax: num('BOARD', 'AM8') || 0,
      goal: num('월말일현황', 'Y2') || ps.goal || 1000000000, tickers: tickers,
      addBuy: ps.addBuy || {}, limits: { pension: pension, isa: isa }, taxSell: ps.taxSell || {}
    };
    /* --- 목표비중 (입력칸만 읽고 계산은 앱에서) --- */
    if (wb.Sheets['목표비중']) {
      var TG = '목표비중', tn = function (a, d) { var x = num(TG, a); return x == null ? d : x; };
      settings.target = {
        buy: tn('B1', 10000000), cashInc: tn('E1', 0), months: tn('A3', 6),
        alloc: { SPCX: tn('B4', 0), NSD100: tn('C4', 0), 'S&P500': tn('D4', 0), SCHD: tn('E4', 0), GOLD: tn('F4', 0) },
        rise: { SPCX: tn('C12', 0), NSD100: tn('D12', 0), 'S&P500': tn('E12', 0), SCHD: tn('F12', 0), GOLD: tn('G12', 0), CASH: tn('H12', 0) }
      };
    } else if (ps.target) settings.target = ps.target;

    /* --- 거래내역 --- */
    var trades = [], last = lastRow('거래내역');
    for (r = 4; r <= last; r++) {
      var t = str('거래내역', 'C' + r); if (!t) continue;
      var gc = c('거래내역', 'G' + r), usd = !!(gc && (gc.v != null || gc.f));
      var row = { id: 'r' + String(r).padStart(4, '0'), t: t, p: r6(num('거래내역', 'D' + r)), q: r6(num('거래내역', 'E' + r)) };
      var d = iso('거래내역', 'B' + r); if (d) row.d = d;
      if (row.p == null || row.q == null) continue;
      if (usd) {
        var gf = f('거래내역', 'G' + r);
        if (gf) { var m = gf.match(/BOARD!(\$[A-Z]+\$\d+)/); if (!m || refMap[m[1]] !== t) row.fx = r6(num('거래내역', 'G' + r)); }
        else row.fx = r6(num('거래내역', 'G' + r));
        var hf = f('거래내역', 'H' + r);
        if (!hf && num('거래내역', 'H' + r) != null) row.krw = r6(num('거래내역', 'H' + r));
        if (!tickers.some(function (x) { return x.t === t; })) row.c = 'USD';
      } else if (!tickers.some(function (x) { return x.t === t; })) row.c = 'KRW';
      var memo = str('거래내역', 'I' + r); if (memo) row.m = memo;
      var sd = iso('거래내역', 'J' + r);
      if (sd || num('거래내역', 'K' + r) != null) {
        row.sd = sd; row.sq = r6(num('거래내역', 'K' + r)); row.sp = r6(num('거래내역', 'L' + r));
        var mf = f('거래내역', 'M' + r), mm = mf && mf.match(/\*\s*([\d.]+)\s*$/);
        if (usd && mm) row.sfx = +mm[1];
      }
      trades.push(row);
    }

    /* --- 월말일현황 --- */
    var mrows = [], liveMemo = {}, liveIssue = {}, yr = null, lastM = lastRow('월말일현황');
    for (r = 5; r <= lastM; r++) {
      var a = v('월말일현황', 'A' + r); if (typeof a === 'number' && a > 1900 && a < 2200 && !f('월말일현황', 'A' + r)) yr = a;
      var mo = num('월말일현황', 'B' + r); if (!yr || !mo) continue;
      var ym = yr + '-' + String(mo).padStart(2, '0');
      var jf = f('월말일현황', 'J' + r), jv = num('월말일현황', 'J' + r);
      var issue = [str('월말일현황', 'C' + r), str('월말일현황', 'D' + r)].filter(Boolean).join(' / '), memo2 = str('월말일현황', 'O' + r);
      if (jf) { if (jf.indexOf('$J$2') < 0) { if (memo2) liveMemo[ym] = memo2; if (issue) liveIssue[ym] = issue; } continue; }
      if (jv == null) continue;
      mrows.push({ ym: ym, issue: issue, fx: r6(num('월말일현황', 'E' + r)), buy: r6(num('월말일현황', 'G' + r)) || 0, total: r6(jv), stock: r6(num('월말일현황', 'K' + r)) || 0, cash: r6(num('월말일현황', 'M' + r)) || 0, memo: memo2 });
    }
    var ms = [];
    ['AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ'].forEach(function (col) { var lab = str('월말일현황', col + '1'); if (lab) ms.push({ label: lab, date: iso('월말일현황', col + '2') }); });
    var monthend = { rows: mrows, liveMemo: liveMemo, liveIssue: liveIssue, milestones: ms };

    /* --- 배당 --- */
    var cols = 'DEFGHIJKLMNOPQ'.split(''), defs = [];
    for (var i = 0; i < cols.length; i++) {
      var nm = str('배당', cols[i] + '6'); if (!nm) continue;
      if (/USD/.test(str('배당', cols[i] + '7'))) defs.push({ t: nm, u: cols[i], k: cols[i + 1] }); else defs.push({ t: nm, u: null, k: cols[i] });
    }
    var divs = [], lastD = lastRow('배당');
    for (r = 9; r <= lastD; r++) {
      var yy = (str('배당', 'B' + r).match(/(\d{4})/) || [])[1], mm2 = (str('배당', 'C' + r).match(/(\d{1,2})/) || [])[1];
      if (!yy || !mm2) continue;
      defs.forEach(function (dd) {
        var u = dd.u ? num('배당', dd.u + r) : null, k = num('배당', dd.k + r);
        if (u || k) divs.push({ ym: yy + '-' + String(mm2).padStart(2, '0'), t: dd.t, usd: u || null, krw: k || 0 });
      });
    }
    var pd = (prev && prev.dividend) || {};
    var dividend = { rows: divs, retire: { cash: num('배당', 'AD22') || (pd.retire || {}).cash || 0, rate: num('배당', 'AE22') || (pd.retire || {}).rate || 0 } };

    /* --- 자녀증여 --- */
    function kid(nameA, birthA, cashA, dc, bc) {
      var nmx = str('자녀증여', nameA), deps = [], buys = [], lk = lastRow('자녀증여');
      for (var rr = 10; rr <= lk; rr++) {
        var d1 = iso('자녀증여', dc[0] + rr); if (d1) deps.push({ d: d1, amt: num('자녀증여', dc[1] + rr) || 0, m: str('자녀증여', dc[2] + rr) });
        var d2 = iso('자녀증여', bc[0] + rr); if (d2 && str('자녀증여', bc[1] + rr)) buys.push({ d: d2, t: str('자녀증여', bc[1] + rr), p: num('자녀증여', bc[2] + rr) || 0, q: num('자녀증여', bc[3] + rr) || 0 });
      }
      var mt = nmx.match(/^(.*?)\s*\((.*)\)\s*$/);
      return { name: mt ? mt[1] : nmx, acct: mt ? mt[2] : '', birth: iso('자녀증여', birthA), cash: num('자녀증여', cashA) || 0, deps: deps, buys: buys };
    }
    var pk = (prev && prev.kids) || {};
    var kids = {
      list: [kid('B2', 'C3', 'F3', ['B', 'C', 'D'], ['F', 'G', 'H', 'I']), kid('M2', 'N3', 'Q3', ['M', 'N', 'O'], ['Q', 'R', 'S', 'T'])],
      est: { rate: (pk.est || {}).rate || 0.1, monthly: (pk.est || {}).monthly || 150000, y1: num('자녀증여', 'C4') || 10, y2: num('자녀증여', 'H4') || 20 },
      priceMap: pk.priceMap || { 'TIGER S&P500': tickers[4].t, 'TIGER나스닥100': tickers[3].t }
    };
    kids.priceMap[str('자녀증여', 'C6') || 'TIGER S&P500'] = tickers[4].t;
    kids.priceMap[str('자녀증여', 'C7') || 'TIGER나스닥100'] = tickers[3].t;

    /* --- PLAN --- */
    var pp = (prev && prev.plan) || {}, RISK = { SCHD: '중립', QQQM: '고위험', 'S&P500': '고위험', '고배당커버드': '중립', '채권': '안전', '예금': '안전' };
    var hasE = !!wb.Sheets['2031-40추계'];
    var alloc = ['C', 'D', 'E', 'F', 'G', 'H'].map(function (col, i) {
      var name = str('PLAN', col + '13'), old = (pp.alloc || [])[i] || {};
      return { name: name, w: num('PLAN', col + '14') || 0, dy: num('PLAN', col + '16') || 0, gr: hasE ? (num('2031-40추계', 'C' + (9 + i)) || 0) : (old.gr || 0), risk: old.risk || RISK[name] || '중립' };
    });
    var plan = {
      target: iso('PLAN', 'C2'), tax: num('PLAN', 'M3') || 0.15,
      people: [['B3', 'B4'], ['D3', 'D4'], ['F3', 'F4', 1], ['H3', 'H4', 1]].map(function (x) { var o = { name: str('PLAN', x[0]), birth: iso('PLAN', x[1]) }; if (x[2]) o.kid = true; return o; }),
      aptSale: num('PLAN', 'B9') || 0, loan: num('PLAN', 'C9') || 0, stock: num('PLAN', 'E9') || 0, cash: num('PLAN', 'F9') || 0, deposit: num('PLAN', 'H9') || 0,
      alloc: alloc,
      income: [{ who: str('PLAN', 'B21'), kind: str('PLAN', 'C21'), amt: null, memo: '배당(자동)' }, { who: str('PLAN', 'B22'), kind: str('PLAN', 'C22'), amt: num('PLAN', 'D22') || 0, memo: str('PLAN', 'E22') }, { who: str('PLAN', 'B23'), kind: str('PLAN', 'C23'), amt: num('PLAN', 'D23') || 0, memo: str('PLAN', 'E23') }],
      expense: [21, 22, 23].map(function (rr) { return { kind: str('PLAN', 'F' + rr), desc: str('PLAN', 'G' + rr), amt: num('PLAN', 'H' + rr) || 0 }; })
    };

    /* --- SIM --- */
    var sim = { fxNow: num('SIM', 'B8'), prices: {} };
    POS.forEach(function (p, i) { sim.prices[tickers[i].t] = num('SIM', p.price); });

    return { settings: settings, trades: trades, sim: sim, monthend: monthend, dividend: dividend, kids: kids, plan: plan, source: 'excel', check: { total: num('BOARD', 'X3') } };
  }
  return { parse: parse };
})();
if (typeof module !== 'undefined') module.exports = XIMP;
