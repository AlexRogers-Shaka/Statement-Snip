// MAPPER-START
// Matches snipped statement lines to rows on the servicing workbook's party tabs.
// Layout must match build2.py.
const LAYOUT = {structFirst:7, n:15, servFirst:6, imFirst:59, imBlock:12, liabFirst:7, liabLast:86,
  systemSheets:['Cover','Guide','Assumptions','Structure','Entity template','Person template','Liabilities','Servicing']};

const TARGETS = {
  pl: [[7,'Revenue'],[8,'Cost of goods sold'],[10,'Other income'],[11,'Operating expenses'],[13,'Depreciation & amortisation'],
       [15,'Interest expense'],[17,'Reported net profit before tax'],[25,'Income tax paid']],
  bs: [[30,'Cash'],[31,'Trade debtors'],[32,'Inventory / WIP'],[33,'Other current assets'],[35,'Property, plant & equipment'],
       [36,'Loans to related parties'],[37,'Other non-current assets'],[39,'Trade creditors'],[40,'Other current liabilities'],
       [41,'Current borrowings'],[43,'Non-current borrowings'],[44,'Loans from related parties'],[45,'Other non-current liabilities']],
  tax: [[7,'Item 1 Salary or wages'],[8,'Item 2 Allowances, director\'s fees'],[9,'Item 10 Gross interest'],[10,'Item 11 Dividends'],
        [11,'Item 13 Distributions'],[13,'Item 15/P8 Business income'],[14,'Item 21 Gross rent'],[15,'Item 21 Rental interest'],
        [16,'Item 21 Capital works'],[17,'Item 21 Other rental deductions'],[19,'Other income'],[20,'Taxable income'],[21,'Tax payable / Medicare levy']],
};
const COST_ROWS = new Set([8,11,13,15,25]);
const INCOME_ROWS = new Set([7,10,17]);
const LIAB_ROWS = new Set([39,40,41,43,44,45]);

const RX = {
  total: /^(total|gross profit|net (profit|loss|income|surplus|deficit|assets|liabilities)|ebitda|ebit\b|operating profit|profit (before|after)|(profit|loss) for the|total comprehensive)/i,
  npbt: /^(net )?(operating )?(profit|loss|surplus|income)( ?\/ ?\(?(loss|deficit)\)?)?\s*(before|pre)[- ]?(income )?tax|^net (profit|loss|surplus|income)( ?\/ ?\(?(loss|deficit)\)?)?$/i,
  after: /after (income )?tax/i,
};

function plSection(label){
  const s = label.toLowerCase();
  if (/cost of (sales|goods)|direct costs|^less:? ?cost/.test(s)) return 'cogs';
  if (/other income|non[- ]operating income/.test(s)) return 'other';
  if (/expense|overhead|outgoings/.test(s)) return 'expense';
  if (/^(trading |operating )?(income|revenue|sales|turnover)\b/.test(s)) return 'income';
  return null;
}
function mapPL(label, section){
  const s = label.toLowerCase();
  if (/income tax|tax expense/.test(s)) return {row:25};
  if (/depreciation|amorti[sz]/.test(s)) return {row:13};
  if (/interest|borrowing costs|finance (costs|charges)/.test(s)){
    if (section==='income' || section==='other' || /received|income|earned|revenue/.test(s)) return {row:10};
    return {row:15};
  }
  if (section==='cogs' || /cost of (sales|goods)|purchases|opening (stock|inventory)|closing (stock|inventory)|direct (costs?|labou?r|wages|materials)|subcontract/.test(s)){
    return {row:8, neg: /closing (stock|inventory)/.test(s)};
  }
  if (section==='income'){
    if (/other income|dividend|gain|profit on (sale|disposal)|grant|government|subsidy|jobkeeper|cash ?flow boost|insurance (recover|claim|proceeds)|rent(al)? (received|income)|sundry|discount received/.test(s)) return {row:10};
    return {row:7};
  }
  if (section==='other') return {row:10};
  if (section==='expense') return {row:11};
  if (/sales|revenue|turnover|fees (income|received)|contract income|commission income/.test(s)) return {row:7};
  if (/other income|dividend|grant|gain on/.test(s)) return {row:10};
  return {row:null, why:'No section heading found – choose a row'};
}

function bsSection(label){
  const s = label.toLowerCase();
  if (/non[- ]?current assets|fixed assets|long[- ]term assets/.test(s)) return 'nca';
  if (/current assets/.test(s)) return 'ca';
  if (/non[- ]?current liabilities|long[- ]term liabilities/.test(s)) return 'ncl';
  if (/current liabilities/.test(s)) return 'cl';
  if (/equity|shareholders.? funds|beneficiar(y|ies).? funds|members.? funds|partners.? funds|capital account|^capital$/.test(s)) return 'eq';
  if (/^assets$/.test(s)) return 'a';
  if (/^liabilities$/.test(s)) return 'l';
  return null;
}
const RELATED = /(loan|advance|account|amount).*(director|shareholder|related|beneficiar|unitholder|associat|partner|member|trust\b|family|div(ision)? ?7a)|(director|shareholder|beneficiar|unitholder|related part|associated entit).*(loan|account|advance)|unpaid present entitlement|\bupe\b/i;
const BORROW = /loan|borrowing|hire purchase|chattel|equipment finance|lease liabilit|finance lease|overdraft|credit card|mortgage|bank bill|facility|line of credit|insurance premium funding/i;
const TAXPROV = /gst|\bbas\b|payg|\bato\b|tax|superannuation|super payable|accru|provision|wages payable|deposit|unearned|in advance|deferred revenue|annual leave|long service/i;
const PPE = /plant|equipment|vehicle|motor|land|building|property|furniture|fixtures|fittings|leasehold|improvement|machinery|computer|office equipment|tools|right[- ]of[- ]use|accumulated depreciation|less:? ?(accumulated )?depreciation|written down|at cost/i;
function mapBS(label, section){
  const s = label.toLowerCase();
  if (section==='eq') return {row:null, skip:true, why:'Equity – not needed'};
  const asset = section==='ca' || section==='nca' || section==='a';
  const liab = section==='cl' || section==='ncl' || section==='l';
  if (asset){
    if (RELATED.test(s)) return {row:36};
    if (/cash|bank|cheque|savings|term deposit|petty|at call/.test(s) && !/loan/.test(s)) return {row:30};
    if (/debtor|receivable/.test(s)) return {row:31};
    if (/stock|inventor|work in progress|\bwip\b/.test(s)) return {row:32};
    if (PPE.test(s)) return {row:35, neg: /accumulated|less|provision for depreciation/.test(s)};
    if (section==='ca') return {row:33};
    if (section==='nca') return {row:37};
    return {row:null, why:'Choose current or non-current'};
  }
  if (liab){
    if (RELATED.test(s)) return {row:44};
    if (TAXPROV.test(s)) return {row: section==='ncl' ? 45 : 40};
    if (/creditor|payable|trade accounts/.test(s)) return {row:39};
    if (BORROW.test(s)) return {row: section==='cl' ? 41 : 43};
    if (section==='cl') return {row:40};
    if (section==='ncl') return {row:45};
    return {row:null, why:'Choose current or non-current'};
  }
  return {row:null, why:'No section heading found – choose a row'};
}

function mapTax(label){
  const s = label.toLowerCase();
  if (/franking credit|tax withheld|tax offset|\bpayg\b.*(withheld|instal)/.test(s)) return {row:null, skip:true, why:'Credit/offset – not income'};
  if (/^total|total (income|deductions)|^deductions|work[- ]related|gifts|donations|self[- ]education|cost of managing/.test(s)) return {row:null, skip:true, why:'Total or deduction – not used'};
  if (/tax payable|tax on taxable income|net tax|medicare levy|total tax/.test(s) && !/surcharge exemption/.test(s)) return {row:21};
  if (/taxable income/.test(s)) return {row:20};
  if (/capital works/.test(s)) return {row:16};
  if (/(rental|rent).*interest|interest.*(deduct|rental|investment property)/.test(s)) return {row:15};
  if (/(other )?rental deduction|rental expense/.test(s)) return {row:17};
  if (/gross rent|rental income|^rent\b|rent received/.test(s)) return {row:14};
  if (/salary|wages/.test(s) && !/allowance/.test(s)) return {row:7};
  if (/allowance|earnings|tips|director.?s? fees?|commissions?/.test(s)) return {row:8};
  if (/distribution|partnership|trust income|share of net income/.test(s)) return {row:11};
  if (/business income|income or loss from business|net (income|loss) from business|sole trader|^p8\b|net business/.test(s)) return {row:13};
  if (/dividend/.test(s)) return {row:10};
  if (/interest/.test(s)) return {row:9};
  if (/capital gain/.test(s)) return {row:null, why:'Capital gain – usually excluded; choose Other income to include'};
  if (/foreign|annuit|pension|other income|employment termination|bonus/.test(s)) return {row:19};
  return {row:null, why:'Not recognised – choose a row'};
}

function guessStatement(rows){
  const text = rows.map(r=>String(r.cells[0]||'').toLowerCase()).join(' | ');
  const score = (re)=> (text.match(re)||[]).length;
  const bs = score(/current assets|current liabilities|total assets|total liabilities|net assets|equity|debtors|creditors|receivable|payable/g);
  const pl = score(/revenue|sales|income|expenses|cost of sales|gross profit|net profit|depreciation|wages/g);
  const tx = score(/salary or wages|taxable income|item \d|allowances|franking|medicare|gross rent|tax payable/g);
  if (tx >= 2 && tx >= bs) return 'tax';
  return bs > pl ? 'bs' : 'pl';
}

// rows: [{cells:[label,...], type}] ; returns [{row, neg, skip, why, auto}]
function autoMap(rows, kind){
  let section = null; let npbtFound = false;
  return rows.map(r=>{
    const label = String(r.cells[0]||'').trim();
    const vals = r.cells.slice(1).map(v=>String(v||'').trim()).filter(Boolean);
    const hasVals = vals.length > 0;
    if (hasVals && vals.every(v=>/^(FY|CY)?(19|20)\d{2}$/i.test(v))) return {row:null, skip:true, why:'Heading'};
    if (kind!=='tax'){
      const sec = kind==='pl' ? plSection(label) : bsSection(label);
      if (sec && (!hasVals || r.type==='heading')) { section = sec; return {row:null, skip:true, why:'Heading'}; }
      if (!hasVals) return {row:null, skip:true, why:'No figures'};
      if (kind==='pl' && RX.npbt.test(label) && !RX.after.test(label) && !npbtFound){ npbtFound = true; return {row:17, why:'Used to check nothing was missed'}; }
      if (RX.total.test(label) || r.type==='total' || /^total\b/i.test(label)) {
        if (kind==='bs' && sec) section = sec;
        return {row:null, skip:true, why:'Total – skipped so nothing is counted twice'};
      }
      if (sec) section = sec;
      return kind==='pl' ? mapPL(label, section) : mapBS(label, section);
    }
    if (!hasVals) return {row:null, skip:true, why:'No figures'};
    return mapTax(label);
  });
}

// Which FY column (0..2) each snipped value column goes to; -1 = skip.
function hintYear(rows){
  for (const r of rows.slice(0,6)){ const m = r.cells.join(' ').match(/(?:FY|year ended[^0-9]*|30 june |june )?((19|20)\d{2})/i); if (m) return m[1]; }
  return null;
}
function guessYears(cols, fyLabels, rows){
  const hint = rows ? hintYear(rows) : null;
  if (hint && cols.length === 2 && !/(19|20)\d{2}/.test(String(cols[1]))) cols = [cols[0], hint];
  const fyYears = fyLabels.map(l=>{ const m = String(l).match(/(19|20)\d{2}/); if (m) return m[0]; const m2 = String(l).match(/\d{2}$/); return m2 ? '20'+m2[0] : null; });
  const out = []; const used = new Set(); let order = [2,1,0];
  cols.slice(1).forEach((h,i)=>{
    const t = String(h||'');
    if (/^note/i.test(t)) { out.push(-1); return; }
    let m = t.match(/(19|20)\d{2}/); let y = m ? m[0] : null;
    if (!y){ const m2 = t.match(/^(FY|CY|YE)?\s*['’]?(\d{2})$/i); if (m2) y = '20'+m2[2]; }
    const idx = y ? fyYears.indexOf(y) : -1;
    if (idx >= 0 && !used.has(idx)){ out.push(idx); used.add(idx); return; }
    out.push(null);
  });
  // fill unknowns in order latest → oldest (statements usually show current year first)
  for (let i=0;i<out.length;i++){
    if (out[i]===null){ const nx = order.find(k=>!used.has(k)); if (nx===undefined) out[i]=-1; else { out[i]=nx; used.add(nx); } }
  }
  return out;
}

// Combine mapped rows into {targetRow: [fy0, fy1, fy2]} applying sign rules.
function aggregate(rows, maps, yearCols, toNum){
  const vals = {}; // row -> [a,b,c]
  const add = (row, fy, v)=>{ if (!vals[row]) vals[row] = [null,null,null]; vals[row][fy] = (vals[row][fy]||0) + v; };
  // detect negative convention for costs / liabilities
  let negCost = 0, posCost = 0;
  rows.forEach((r,i)=>{ const m = maps[i]; if (!m || m.row==null || !(COST_ROWS.has(m.row) || LIAB_ROWS.has(m.row)) || m.neg) return;
    r.cells.slice(1).forEach((c,j)=>{ if (yearCols[j] < 0) return; const n = toNum(c); if (n==null) return; if (n<0) negCost++; else if (n>0) posCost++; }); });
  const flipCost = negCost > posCost;
  rows.forEach((r,i)=>{
    const m = maps[i]; if (!m || m.row==null) return;
    r.cells.slice(1).forEach((c,j)=>{
      const fy = yearCols[j]; if (fy==null || fy<0) return;
      let n = toNum(c); if (n==null) return;
      if (flipCost && (COST_ROWS.has(m.row) || LIAB_ROWS.has(m.row))) n = -n;
      if (m.neg) n = -Math.abs(n);
      add(m.row, fy, n);
    });
  });
  return vals;
}
// MAPPER-END
if (typeof module !== 'undefined') module.exports = {autoMap, guessStatement, guessYears, aggregate, TARGETS, LAYOUT};
