

/* ===== MV Family Office - parse-core =====
   Pure parsing logic (no DOM). Works with SheetJS workbook object
   (same shape in Node 'xlsx' package and browser XLSX global).
   Exposed as window.MVParse in the browser (see bottom).
*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.MVParse = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {

  // ---------- month name lookup (Spanish + English, incl. known typo) ----------
  const MONTHS = {
    'enero':0,'febrero':1,'marzo':2,'abril':3,'mayo':4,'junio':5,'julio':6,'agosto':7,
    'septiembre':8,'septembre':8,'setiembre':8,'octubre':9,'noviembre':10,'diciembre':11,
    'jan':0,'feb':1,'mar':2,'apr':3,'may':4,'jun':5,'jul':6,'aug':7,'sep':8,'sept':8,'oct':9,'nov':10,'dec':11
  };
  const SKIP_TOKENS = new Set(['ytd','totals','total','%']);

  // ---------- low level cell helpers ----------
  function colToLetter(c) {
    let s = '';
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function letterToCol(l) {
    let c = 0;
    for (let i = 0; i < l.length; i++) c = c * 26 + (l.charCodeAt(i) - 64);
    return c;
  }
  function cellRef(row, col) { return colToLetter(col) + row; }
  function getCell(ws, row, col) {
    if (!ws) return undefined;
    const ref = cellRef(row, col);
    const c = ws[ref];
    if (!c) return undefined;
    if (c.t === 'e') return undefined; // Excel error cell (#DIV/0!, #REF!, etc.)
    return c.v;
  }
  function sheetBounds(ws) {
    if (!ws || !ws['!ref']) return { minRow: 1, maxRow: 1, minCol: 1, maxCol: 1 };
    const range = ws['!ref'].split(':');
    const parse = (ref) => {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      return { col: letterToCol(m[1]), row: parseInt(m[2], 10) };
    };
    const a = parse(range[0]), b = parse(range[1] || range[0]);
    return { minRow: a.row, maxRow: b.row, minCol: a.col, maxCol: b.col };
  }
  function normLabel(v) {
    if (v === undefined || v === null) return '';
    return String(v).replace(/\s+/g, ' ').trim().toLowerCase();
  }
  // Find first row (within [rowStart,rowEnd]) whose value in colIdx exactly matches label (normalized)
  function findRowByExactLabel(ws, colIdx, label, rowStart, rowEnd) {
    const target = normLabel(label);
    for (let r = rowStart; r <= rowEnd; r++) {
      if (normLabel(getCell(ws, r, colIdx)) === target) return r;
    }
    return null;
  }
  function findRowByLabelContains(ws, colIdx, needle, rowStart, rowEnd) {
    const target = normLabel(needle);
    for (let r = rowStart; r <= rowEnd; r++) {
      const v = normLabel(getCell(ws, r, colIdx));
      if (v && v.indexOf(target) !== -1) return r;
    }
    return null;
  }

  function monthIndexFromValue(v) {
    if (v instanceof Date) return { year: v.getFullYear(), month: v.getMonth() };
    if (typeof v === 'string') {
      const key = normLabel(v);
      if (SKIP_TOKENS.has(key)) return null;
      if (key in MONTHS) return { year: null, month: MONTHS[key] };
    }
    return null;
  }

  // Build a map colIndex -> {year, month} by forward-filling a sparse "year" row
  // and reading a "month name / date" row for every column.
  function buildDateColMap(ws, yearRow, monthRow, colStart, colEnd) {
    const map = {};
    let currentYear = null;
    for (let c = colStart; c <= colEnd; c++) {
      const yv = getCell(ws, yearRow, c);
      if (typeof yv === 'number' && yv > 1900 && yv < 2100) currentYear = yv;
      const mv = getCell(ws, monthRow, c);
      const parsed = monthIndexFromValue(mv);
      if (parsed) {
        const year = parsed.year !== null ? parsed.year : currentYear;
        if (year !== null) map[c] = { year, month: parsed.month };
      }
    }
    return map;
  }
  function findColForYearMonth(map, year, month) {
    for (const c in map) {
      if (map[c].year === year && map[c].month === month) return parseInt(c, 10);
    }
    return null;
  }
  function lastColWithYearMonth(map) {
    let best = null, bestKey = -1;
    for (const c in map) {
      const key = map[c].year * 12 + map[c].month;
      if (key > bestKey) { bestKey = key; best = parseInt(c, 10); }
    }
    return best;
  }

  function isNum(v) { return typeof v === 'number' && !isNaN(v); }
  function safe(v, fallback) { return isNum(v) ? v : (fallback !== undefined ? fallback : null); }

  // ===================================================================
  // SECTION PARSERS
  // ===================================================================

  function parseMeta(wb) {
    const banks = wb.Sheets['Banks'];
    const asOf = getCell(banks, 4, 6); // F4
    const priorYearEnd = getCell(banks, 3, 6); // F3
    const eurUsd = safe(getCell(banks, 3, 16), null); // P3
    const sofr = safe(getCell(banks, 4, 16), null); // P4
    return { asOf: asOf instanceof Date ? asOf : null, priorYearEnd: priorYearEnd instanceof Date ? priorYearEnd : null, eurUsd, sofr };
  }

  function parseDashboard(wb) {
    const ws = wb.Sheets['Dashboard'];
    if (!ws) return null;
    return {
      totalAssets: safe(getCell(ws, 16, 11)),   // K16
      totalLiabilities: safe(getCell(ws, 26, 11)), // K26
      netWorth: safe(getCell(ws, 21, 15)),      // O21
      banksTotal: safe(getCell(ws, 11, 8)),     // H11
      realEstate: safe(getCell(ws, 16, 8)),     // H16
      artOthers: safe(getCell(ws, 22, 8)),      // H22
      allocation: {
        cash: safe(getCell(ws, 10, 21)),        // U10
        fixedIncome: safe(getCell(ws, 10, 23)), // W10
        equity: safe(getCell(ws, 10, 25)),      // Y10
        alternative: safe(getCell(ws, 10, 27)), // AA10
        realEstate: safe(getCell(ws, 10, 29)),  // AC10
        art: safe(getCell(ws, 10, 31))          // AE10
      }
    };
  }

  const BANK_ROW_LABELS = ['JPM US', 'JPM Suiza', 'UBS Suiza', 'UBS US', 'EdR', 'Citi', 'Venecredit'];
  const BANK_SHEET_MAP = {
    'JPM US': 'JPM(US)', 'JPM Suiza': 'JPM(CH)', 'UBS Suiza': 'UBS(CH)',
    'UBS US': 'UBS(US)', 'EdR': 'EdR', 'Citi': 'CITI', 'Venecredit': 'VENECREDIT'
  };
  const LOAN_ROW_LABELS = { 'JPM US': 'JPM US', 'UBS Suiza': 'UBS Suiza', 'UBS US': 'UBS US', 'EdR': 'EdR' };

  function parseBanks(wb) {
    const ws = wb.Sheets['Banks'];
    if (!ws) return { rows: [], total: null };
    const rows = [];
    for (let r = 10; r <= 16; r++) {
      const name = getCell(ws, r, 4); // D
      if (!name) continue;
      rows.push({
        name: String(name).trim(),
        currency: getCell(ws, r, 5) || 'USD', // E
        balancePriorYearEnd: safe(getCell(ws, r, 6), 0), // F
        balanceLastMonth: safe(getCell(ws, r, 7), 0),    // G
        balance: safe(getCell(ws, r, 8), 0),             // H
        monthReturn: safe(getCell(ws, r, 12)),           // L
        ytdReturn: safe(getCell(ws, r, 17))              // Q
      });
    }
    return {
      rows,
      total: {
        balance: safe(getCell(ws, 17, 8)),      // H17
        monthReturnGross: safe(getCell(ws, 17, 12)), // L17
        monthReturnNet: safe(getCell(ws, 21, 12)),   // L21
        ytdReturnGross: safe(getCell(ws, 17, 16)),   // P17
        ytdReturnNet: safe(getCell(ws, 17, 17))      // Q17
      }
    };
  }

  function parseLoan(wb) {
    const ws = wb.Sheets['Loan'];
    if (!ws) return { rows: [], total: null };
    const rows = [];
    for (let r = 4; r <= 7; r++) {
      const name = getCell(ws, r, 2); // B
      if (!name) continue;
      rows.push({
        name: String(name).trim(),
        balance: safe(getCell(ws, r, 3), 0),      // C
        lineUsed: safe(getCell(ws, r, 4), 0),      // D
        netBalance: safe(getCell(ws, r, 5), 0),    // E
        lineAvailable: safe(getCell(ws, r, 6), 0), // F
        usedPct: safe(getCell(ws, r, 7), 0),       // G
        notUsedPct: safe(getCell(ws, r, 8), 0)     // H
      });
    }
    return {
      rows,
      total: {
        balance: safe(getCell(ws, 8, 3)),
        lineUsed: safe(getCell(ws, 8, 4)),
        netBalance: safe(getCell(ws, 8, 5)),
        lineAvailable: safe(getCell(ws, 8, 6)),
        usedPct: safe(getCell(ws, 8, 7))
      }
    };
  }

  // Historical annual: returns {gross:[{year,finalBalance,perf}], net:[...]}
  function parseHistoricalAnnual(wb) {
    const ws = wb.Sheets['Historical'];
    if (!ws) return { gross: [], net: [] };
    const bounds = sheetBounds(ws);
    const grossHeaderRow = findRowByExactLabel(ws, 4, 'Initial Balance', 1, 30); // col D
    const netHeaderRow = findRowByExactLabel(ws, 4, 'Initial Balance', (grossHeaderRow || 1) + 1, bounds.maxRow);
    function readBlock(headerRow, endRow) {
      if (!headerRow) return [];
      const out = [];
      for (let r = headerRow + 1; r <= endRow; r++) {
        const year = getCell(ws, r, 3); // C
        if (typeof year !== 'number') continue;
        const finalBalance = safe(getCell(ws, r, 12)); // L
        const perf = safe(getCell(ws, r, 13));         // M
        if (finalBalance === null) continue;
        out.push({ year, finalBalance, perf });
        if (year > 2100) break;
      }
      return out;
    }
    return {
      gross: readBlock(grossHeaderRow, netHeaderRow ? netHeaderRow - 1 : bounds.maxRow),
      net: readBlock(netHeaderRow, bounds.maxRow)
    };
  }

  // Net Balances monthly series (financial net worth)
  function parseNetBalancesMonthly(wb) {
    const ws = wb.Sheets['Net Balances'];
    if (!ws) return [];
    const bounds = sheetBounds(ws);
    const out = [];
    for (let r = 4; r <= bounds.maxRow; r++) {
      const date = getCell(ws, r, 3); // C
      if (!(date instanceof Date)) continue;
      const balanceInBanks = safe(getCell(ws, r, 4), 0); // D
      const totalLoan = safe(getCell(ws, r, 8), 0);       // H
      const realEstate = safe(getCell(ws, r, 9), 0);      // I
      const netWorth = safe(getCell(ws, r, 11), 0);        // K
      if (balanceInBanks === 0 && out.length > 0) break;
      out.push({ date, balanceInBanks, totalLoan, realEstate, netWorth, netWorthWithRE: balanceInBanks + totalLoan + realEstate });
    }
    return out;
  }

  // Performance sheet: monthly CEMS Gross/Net returns
  function parsePerformanceMonthly(wb) {
    const ws = wb.Sheets['Performance'];
    if (!ws) return [];
    const bounds = sheetBounds(ws);
    const rowGross = findRowByExactLabel(ws, 7, 'CEMS (Gross)', 1, 40); // col G
    const rowNet = findRowByExactLabel(ws, 7, 'CEMS (Net)', 1, 40);
    if (!rowGross && !rowNet) return [];
    const dateMap = buildDateColMap(ws, 2, 3, 8, bounds.maxCol); // year row2, month row3, data starts col H(8)
    const cols = Object.keys(dateMap).map(Number).sort((a, b) => a - b);
    const out = [];
    for (const c of cols) {
      const g = rowGross ? safe(getCell(ws, rowGross, c)) : null;
      const n = rowNet ? safe(getCell(ws, rowNet, c)) : null;
      if (g === null && n === null) continue;
      const { year, month } = dateMap[c];
      out.push({ year, month, gross: g, net: n });
    }
    return out;
  }

  function findLabelRowNormalized(ws, rowStart, rowEnd, predicate) {
    for (let r = rowStart; r <= rowEnd; r++) {
      const v = normLabel(getCell(ws, r, 2));
      if (v && predicate(v)) return r;
    }
    return null;
  }

  // Per-bank sheet detail: current snapshot (allocation + fees)
  function parseBankDetail(wb, sheetName, asOfDate) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !asOfDate) return null;
    const bounds = sheetBounds(ws);
    const totalsRow = findRowByExactLabel(ws, 2, 'TOTALS', 1, bounds.maxRow);
    if (!totalsRow) return null;
    const dateMap = buildDateColMap(ws, 3, 4, 3, bounds.maxCol);
    const col = findColForYearMonth(dateMap, asOfDate.getFullYear(), asOfDate.getMonth());
    if (!col) return null;

    const cashRow = findRowByExactLabel(ws, 2, 'Cash & Short Term', totalsRow, totalsRow + 25);
    const fiRow = findRowByExactLabel(ws, 2, 'Fixed Income', totalsRow, totalsRow + 25);
    const eqRow = findRowByExactLabel(ws, 2, 'Equity', totalsRow, totalsRow + 25);
    const altRow = findRowByExactLabel(ws, 2, 'Alternative Assets', totalsRow, totalsRow + 25);
    const feesYtdRowIdx = findLabelRowNormalized(ws, totalsRow, bounds.maxRow, (s) => s.indexOf('total fees') !== -1 && s.indexOf('ytd') !== -1);
    const feesMonthRowIdx = findLabelRowNormalized(ws, totalsRow, bounds.maxRow, (s) => s.indexOf('total fees') !== -1 && s.indexOf('month') !== -1);

    return {
      balance: safe(getCell(ws, totalsRow, col)),
      allocation: {
        cash: cashRow ? safe(getCell(ws, cashRow, col), 0) : 0,
        fixedIncome: fiRow ? safe(getCell(ws, fiRow, col), 0) : 0,
        equity: eqRow ? safe(getCell(ws, eqRow, col), 0) : 0,
        alternative: altRow ? safe(getCell(ws, altRow, col), 0) : 0
      },
      feesMonth: feesMonthRowIdx ? safe(getCell(ws, feesMonthRowIdx, col), 0) : 0,
      feesYtd: feesYtdRowIdx ? safe(getCell(ws, feesYtdRowIdx, col), 0) : 0
    };
  }

  // Full monthly time series for a single bank sheet: balance, monthly/YTD return,
  // asset allocation and fees for EVERY month present in the sheet (not just "asOf").
  // This is what powers the date-picker "snapshot" feature.
  function parseBankMonthlySeries(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    if (!ws) return [];
    const bounds = sheetBounds(ws);
    const totalsRow = findRowByExactLabel(ws, 2, 'TOTALS', 1, bounds.maxRow);
    if (!totalsRow) return [];
    const dateMap = buildDateColMap(ws, 3, 4, 3, bounds.maxCol);
    const perfRow = findRowByExactLabel(ws, 2, '% Performance', totalsRow, totalsRow + 10);
    const perfYtdRow = findRowByExactLabel(ws, 2, '%Performance YTD', totalsRow, totalsRow + 10);
    const cashRow = findRowByExactLabel(ws, 2, 'Cash & Short Term', totalsRow, totalsRow + 25);
    const fiRow = findRowByExactLabel(ws, 2, 'Fixed Income', totalsRow, totalsRow + 25);
    const eqRow = findRowByExactLabel(ws, 2, 'Equity', totalsRow, totalsRow + 25);
    const altRow = findRowByExactLabel(ws, 2, 'Alternative Assets', totalsRow, totalsRow + 25);
    const feesYtdRowIdx = findLabelRowNormalized(ws, totalsRow, bounds.maxRow, (s) => s.indexOf('total fees') !== -1 && s.indexOf('ytd') !== -1);
    const feesMonthRowIdx = findLabelRowNormalized(ws, totalsRow, bounds.maxRow, (s) => s.indexOf('total fees') !== -1 && s.indexOf('month') !== -1);

    const cols = Object.keys(dateMap).map(Number).sort((a, b) => a - b);
    const out = [];
    let sawNonZero = false;
    for (const c of cols) {
      const balance = safe(getCell(ws, totalsRow, c));
      if (balance === null) continue; // unpopulated cell
      if (balance === 0) {
        if (sawNonZero) break; // was active, this is future padding/inactivation -> stop series here
        continue; // not active yet, skip leading zeros
      }
      sawNonZero = true;
      const { year, month } = dateMap[c];
      out.push({
        year, month, balance,
        monthReturn: perfRow ? safe(getCell(ws, perfRow, c)) : null,
        ytdReturn: perfYtdRow ? safe(getCell(ws, perfYtdRow, c)) : null,
        allocation: {
          cash: cashRow ? safe(getCell(ws, cashRow, c), 0) : 0,
          fixedIncome: fiRow ? safe(getCell(ws, fiRow, c), 0) : 0,
          equity: eqRow ? safe(getCell(ws, eqRow, c), 0) : 0,
          alternative: altRow ? safe(getCell(ws, altRow, c), 0) : 0
        },
        feesMonth: feesMonthRowIdx ? safe(getCell(ws, feesMonthRowIdx, c), 0) : 0,
        feesYtd: feesYtdRowIdx ? safe(getCell(ws, feesYtdRowIdx, c), 0) : 0
      });
    }
    return out;
  }

  // Consolidated monthly time series (all banks + asset allocation) straight from the
  // "Data" master table -- this is what the Dashboard's VLOOKUPs pull from for "this month",
  // generalized here to every month in the file.
  function parseConsolidatedMonthlySeries(wb) {
    const ws = wb.Sheets['Data'];
    if (!ws) return [];
    const bounds = sheetBounds(ws);
    const bankCols = { 'JPM US': 6, 'JPM Suiza': 7, 'UBS Suiza': 8, 'UBS US': 9, 'EdR': 10, 'Citi': 11, 'Venecredit': 12 };
    const out = [];
    for (let r = 4; r <= bounds.maxRow; r++) {
      const date = getCell(ws, r, 5); // E
      if (!(date instanceof Date)) continue;
      const totalBalance = safe(getCell(ws, r, 13), 0); // M
      if (!totalBalance) { if (out.length > 0) break; else continue; }
      const banks = {};
      Object.keys(bankCols).forEach(name => { banks[name] = safe(getCell(ws, r, bankCols[name]), 0); });
      const allocation = {
        cash: safe(getCell(ws, r, 30), 0),       // AD
        fixedIncome: safe(getCell(ws, r, 31), 0), // AE
        equity: safe(getCell(ws, r, 32), 0),      // AF
        alternative: safe(getCell(ws, r, 33), 0)  // AG
      };
      out.push({ year: date.getFullYear(), month: date.getMonth(), date, totalBalance, banks, allocation });
    }
    return out;
  }

  function findFirstRowWithNumericYearAcrossCols(ws, rowStart, rowEnd, maxCol) {
    for (let r = rowStart; r <= rowEnd; r++) {
      for (let c = 5; c <= maxCol; c++) {
        const v = getCell(ws, r, c);
        if (typeof v === 'number' && v > 1900 && v < 2100) return r;
      }
    }
    return null;
  }

  // Expenses: annual summary + full per-year monthly/category breakdown, with
  // "Interest Expense" (cost of servicing the loans) split out from real family spending.
  function parseExpenses(wb) {
    const ws = wb.Sheets['Expenses'];
    if (!ws) return { annual: [], currentYear: null, yearBlocks: [] };
    const bounds = sheetBounds(ws);

    const headerRow = findRowByExactLabel(ws, 2, 'Total Expenses', 1, bounds.maxRow);
    const annual = [];
    if (headerRow) {
      for (let r = headerRow + 1; r <= bounds.maxRow; r++) {
        const year = getCell(ws, r, 1); // A
        if (typeof year !== 'number') break;
        annual.push({
          year,
          totalExpenses: safe(getCell(ws, r, 2), 0),
          gifts: safe(getCell(ws, r, 3), 0),
          expensesMinusGifts: safe(getCell(ws, r, 4), 0)
        });
      }
    }

    const catLabelCol = 4; // D
    const totalsRow = findRowByExactLabel(ws, catLabelCol, 'Totals', 1, bounds.maxRow);
    const interestRow = findRowByExactLabel(ws, catLabelCol, 'Interest Expense', 1, bounds.maxRow);
    const yearBlocks = [];
    if (totalsRow) {
      const yearRow = 3; // stable: each yearly block repeats the same row layout, only columns shift right
      // collect every year block (not just the current one) so loan interest can be split out historically too
      const blocks = [];
      for (let c = 5; c <= bounds.maxCol; c++) {
        const yv = getCell(ws, yearRow, c);
        if (typeof yv === 'number' && yv > 1900 && yv < 2100) blocks.push({ year: yv, blockStart: c });
      }
      blocks.forEach(({ year, blockStart }) => {
        const totalsCol = blockStart + 12;
        const monthly = [];
        const interestMonthly = [];
        for (let i = 0; i < 12; i++) {
          monthly.push(safe(getCell(ws, totalsRow, blockStart + i), 0));
          interestMonthly.push(interestRow ? safe(getCell(ws, interestRow, blockStart + i), 0) : 0);
        }
        const categories = []; // excludes "Interest Expense" -- that's loan cost, not family spending
        for (let r = totalsRow - 14; r < totalsRow; r++) {
          const label = getCell(ws, r, catLabelCol);
          if (!label) continue;
          const labelStr = String(label).trim();
          if (normLabel(labelStr) === 'interest expense') continue;
          const val = safe(getCell(ws, r, totalsCol), 0);
          if (val) categories.push({ label: labelStr, value: val });
        }
        const ytdTotal = safe(getCell(ws, totalsRow, totalsCol), 0) || 0;
        const interestExpense = interestRow ? (safe(getCell(ws, interestRow, totalsCol), 0) || 0) : 0;
        yearBlocks.push({
          year,
          ytdTotal,                                  // includes loan interest (matches annual summary table)
          interestExpense,                           // loan interest / debt-servicing cost, split out
          familyExpenses: ytdTotal - interestExpense, // real family spending, excluding debt cost
          monthly,                                   // includes interest (per-month grand total)
          interestMonthly,                           // interest portion per month
          familyMonthly: monthly.map((m, i) => m - interestMonthly[i]),
          categories                                 // family categories only, interest excluded
        });
      });
    }
    const currentYear = yearBlocks.length ? yearBlocks[yearBlocks.length - 1] : null;
    return { annual, currentYear, yearBlocks };
  }

  function parseAll(wb) {
    const warnings = [];
    const result = {
      meta: null, dashboard: null, banks: null, loan: null, historicalAnnual: null,
      netBalancesMonthly: null, performanceMonthly: null, bankDetails: {}, expenses: null,
      consolidatedMonthlySeries: null, bankMonthlySeries: {}
    };
    function tryStep(name, fn) {
      try { return fn(); } catch (e) { warnings.push(name + ': ' + e.message); return null; }
    }
    result.meta = tryStep('meta', () => parseMeta(wb));
    result.dashboard = tryStep('dashboard', () => parseDashboard(wb));
    result.banks = tryStep('banks', () => parseBanks(wb));
    result.loan = tryStep('loan', () => parseLoan(wb));
    result.historicalAnnual = tryStep('historicalAnnual', () => parseHistoricalAnnual(wb));
    result.netBalancesMonthly = tryStep('netBalancesMonthly', () => parseNetBalancesMonthly(wb));
    result.performanceMonthly = tryStep('performanceMonthly', () => parsePerformanceMonthly(wb));
    result.expenses = tryStep('expenses', () => parseExpenses(wb));
    result.consolidatedMonthlySeries = tryStep('consolidatedMonthlySeries', () => parseConsolidatedMonthlySeries(wb));
    const asOf = result.meta && result.meta.asOf;
    if (result.banks && asOf) {
      for (const row of result.banks.rows) {
        const sheetName = BANK_SHEET_MAP[row.name];
        if (!sheetName) continue;
        result.bankDetails[row.name] = tryStep('bankDetail:' + row.name, () => parseBankDetail(wb, sheetName, asOf));
        result.bankMonthlySeries[row.name] = tryStep('bankMonthlySeries:' + row.name, () => (sheetName ? parseBankMonthlySeries(wb, sheetName) : []));
      }
    }
    result.warnings = warnings;
    return result;
  }

  // ===================================================================
  // DATE-PICKER SUPPORT: list selectable months + build a point-in-time snapshot
  // ===================================================================
  function ymKey(year, month) { return year * 12 + month; }

  function listAvailableDates(result) {
    const series = result.consolidatedMonthlySeries || [];
    return series
      .map(e => ({ year: e.year, month: e.month, key: ymKey(e.year, e.month) }))
      .sort((a, b) => b.key - a.key); // newest first
  }

  function findByYearMonth(arr, year, month, getYM) {
    if (!arr) return null;
    for (const item of arr) {
      const ym = getYM(item);
      if (ym.year === year && ym.month === month) return item;
    }
    return null;
  }

  // Builds a Dashboard/Banks/BankDetails/Loan/Expenses-shaped snapshot for an arbitrary
  // month. When the requested month is the latest one in the file, this returns the exact
  // figures produced by the workbook's own formulas (same as parseDashboard/parseBanks/etc).
  // For earlier months it's reconstructed from the underlying monthly time series -- accurate
  // for balances, allocation and returns; per-bank loan detail is only available for the
  // latest month (flagged via `loanDetailAvailable`), so older snapshots fall back to the
  // aggregate debt figure only.
  function buildSnapshotForDate(result, year, month) {
    const dates = listAvailableDates(result);
    if (!dates.length) return null;
    const latest = dates[0];
    const isLatest = (year === latest.year && month === latest.month);

    if (isLatest) {
      return {
        year, month, isLatest: true, loanDetailAvailable: true,
        dashboard: result.dashboard, banks: result.banks, bankDetails: result.bankDetails,
        loan: result.loan, expensesYear: result.expenses ? result.expenses.currentYear : null
      };
    }

    const cons = findByYearMonth(result.consolidatedMonthlySeries, year, month, e => e);
    const nb = findByYearMonth(result.netBalancesMonthly, year, month, e => ({ year: e.date.getFullYear(), month: e.date.getMonth() }));
    if (!cons) return null;

    const banksTotal = cons.totalBalance;
    const realEstate = nb ? nb.realEstate : (result.dashboard ? result.dashboard.realEstate : 0);
    const totalLoan = nb ? Math.abs(nb.totalLoan) : (result.dashboard ? result.dashboard.totalLiabilities : 0);
    const artOthers = result.dashboard ? result.dashboard.artOthers : 0; // no historical series -- held at latest known value
    const totalAssets = banksTotal + realEstate + artOthers;
    const totalLiabilities = totalLoan;
    const netWorth = totalAssets - totalLiabilities;

    // consolidated monthly/YTD return from the Performance sheet series
    const perfSeries = result.performanceMonthly || [];
    const perfMonth = findByYearMonth(perfSeries, year, month, e => e);
    let ytdGross = null, ytdNet = null;
    const yearMonths = perfSeries.filter(e => e.year === year && e.month <= month).sort((a, b) => a.month - b.month);
    if (yearMonths.length) {
      ytdGross = yearMonths.reduce((acc, e) => acc * (1 + (e.gross || 0)), 1) - 1;
      ytdNet = yearMonths.reduce((acc, e) => acc * (1 + (e.net || 0)), 1) - 1;
    }

    const bankCurrency = {};
    (result.banks ? result.banks.rows : []).forEach(r => { bankCurrency[r.name] = r.currency; });

    const rows = [];
    const bankDetails = {};
    Object.keys(cons.banks).forEach(name => {
      const series = result.bankMonthlySeries && result.bankMonthlySeries[name];
      const entry = findByYearMonth(series, year, month, e => e);
      rows.push({
        name,
        currency: bankCurrency[name] || 'USD',
        balance: cons.banks[name] || 0,
        balanceLastMonth: null,
        balancePriorYearEnd: null,
        monthReturn: entry ? entry.monthReturn : null,
        ytdReturn: entry ? entry.ytdReturn : null
      });
      bankDetails[name] = entry ? {
        balance: entry.balance, allocation: entry.allocation, feesMonth: entry.feesMonth, feesYtd: entry.feesYtd
      } : null;
    });

    const expYear = result.expenses && result.expenses.yearBlocks
      ? result.expenses.yearBlocks.find(b => b.year === year)
      : null;

    return {
      year, month, isLatest: false, loanDetailAvailable: false,
      dashboard: {
        totalAssets, totalLiabilities, netWorth, banksTotal, realEstate, artOthers,
        allocation: {
          cash: cons.allocation.cash, fixedIncome: cons.allocation.fixedIncome,
          equity: cons.allocation.equity, alternative: cons.allocation.alternative,
          realEstate, art: artOthers
        }
      },
      banks: {
        rows,
        total: {
          balance: banksTotal,
          monthReturnGross: perfMonth ? perfMonth.gross : null,
          monthReturnNet: perfMonth ? perfMonth.net : null,
          ytdReturnGross: ytdGross,
          ytdReturnNet: ytdNet
        }
      },
      bankDetails,
      loan: { rows: [], total: { balance: totalLoan, lineUsed: totalLoan, netBalance: null, lineAvailable: null, usedPct: null } },
      expensesYear: expYear || null
    };
  }

  return {
    parseAll, parseMeta, parseDashboard, parseBanks, parseLoan,
    parseHistoricalAnnual, parseNetBalancesMonthly, parsePerformanceMonthly,
    parseBankDetail, parseExpenses, parseBankMonthlySeries, parseConsolidatedMonthlySeries,
    listAvailableDates, buildSnapshotForDate,
    BANK_SHEET_MAP, LOAN_ROW_LABELS,
    _internal: { colToLetter, letterToCol, buildDateColMap, findColForYearMonth }
  };
}));
