// Simulated Excel for testing the Statement Snip panel end to end.
const fs = require('fs');
const { JSDOM } = require('jsdom');
const html = fs.readFileSync(process.argv[2], 'utf8');

function colNum(s){ let n=0; for (const ch of s) n = n*26 + (ch.charCodeAt(0)-64); return n; }
function colStr(n){ let s=''; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; }
function parse(addr){
  if (/^\d+:\d+$/.test(addr)){ const [a,b]=addr.split(':').map(Number); return {rows:true, r0:a, r1:b}; }
  const [a,b] = addr.split(':'); const p = x=>{ const m=x.match(/^([A-Z]+)(\d+)$/); return [colNum(m[1]), +m[2]]; };
  const [c0,r0]=p(a), [c1,r1]=b?p(b):[c0,r0]; return {c0,r0,c1,r1};
}
const book = { sheets: [], log: [] };
function addSheet(name, vis='Visible', cells={}){ const s={name, visibility:vis, cells:{...cells}, hiddenRows:new Set()}; book.sheets.push(s); return s; }
const struct = {};
const parties = [
  ['Trading Co','Company','Family Trust',1,'Borrower','Yes','', 'Trading Co','Entity'],
  ['Director 1','Individual','','','Guarantor','Yes','', 'Director 1','Person'],
];
parties.forEach((p,i)=>{ 'BCDEFGHIJ'.split('').forEach((c,j)=>{ struct[c+(7+i)] = p[j]; }); });
addSheet('Cover'); addSheet('Guide'); addSheet('Assumptions','Visible',{B6:'FY2023',B7:'FY2024',B8:'FY2025'});
addSheet('Structure','Visible',struct); addSheet('Entity template','Hidden',{A2:'(template)'}); addSheet('Person template','Hidden',{A2:'(template)'});
addSheet('Liabilities'); addSheet('Servicing');

function find(name){ return book.sheets.find(s=>s.name.toLowerCase()===String(name).toLowerCase()); }
function wsProxy(s){
  if (!s) return { isNullObject:true };
  return {
    isNullObject:false,
    get name(){ return s.name; }, set name(v){ if (find(v) && find(v)!==s) throw new Error('dup name '+v); s.name=v; },
    get visibility(){ return s.visibility; }, set visibility(v){ s.visibility=v; },
    load(){}, activate(){ book.active=s.name; },
    copy(pos, rel){ if (s.visibility!=='Visible') throw new Error('Cannot copy hidden sheet'); const n={name:s.name+' (2)', visibility:'Visible', cells:{...s.cells}, hiddenRows:new Set()}; const idx=book.sheets.indexOf(find(rel.name)); book.sheets.splice(idx,0,n); return wsProxy(n); },
    getRange(addr){ return rangeProxy(s, addr); },
  };
}
function rangeProxy(s, addr){
  const a = parse(addr);
  return {
    load(){}, address: s.name+'!'+addr,
    get values(){ const out=[]; for(let r=a.r0;r<=a.r1;r++){ const row=[]; for(let c=a.c0;c<=a.c1;c++){ const v=s.cells[colStr(c)+r]; row.push(v===undefined?'':v); } out.push(row);} return out; },
    set values(v){ for(let r=a.r0;r<=a.r1;r++) for(let c=a.c0;c<=a.c1;c++) s.cells[colStr(c)+r]=v[r-a.r0][c-a.c0]; },
    set rowHidden(h){ for(let r=a.r0;r<=a.r1;r++){ if(h) s.hiddenRows.add(r); else s.hiddenRows.delete(r);} },
  };
}
const Excel = { run: async fn => fn({ sync: async()=>{}, workbook: {
  worksheets: { getItemOrNullObject: n=>wsProxy(find(n)), getItem: n=>{ const s=find(n); if(!s) throw new Error('no sheet '+n); return wsProxy(s); },
    load(){}, get items(){ return book.sheets.map(s=>({name:s.name})); }, onActivated:{ add(){} } } } }) };
const Office = { onReady: cb=>setTimeout(()=>cb({host:'Excel'}),0), HostType:{Excel:'Excel'}, context:{ requirements:{ isSetSupported:()=>true } } };

const dom = new JSDOM(html, { runScripts:'dangerously', pretendToBeVisual:true, beforeParse(w){ w.Office=Office; w.Excel=Excel; w.HTMLCanvasElement.prototype.getContext=()=>({}); } });
const w = dom.window, d = w.document;
const tick = (ms=30)=>new Promise(r=>setTimeout(r,ms));
const vis = id => !d.getElementById(id).hidden;
let fails = 0; const ok = (c, m)=>{ console.log((c?'PASS ':'FAIL ')+m); if(!c) fails++; };
(async()=>{
  await tick(100);
  ok(vis('svCard'), 'Servicing workbook card shows');
  ok(/Trading Co/.test(d.getElementById('svParties').textContent) && /Not built/.test(d.getElementById('svParties').textContent), 'Parties listed as not built');
  d.getElementById('svBuild').click(); await tick(100);
  ok(!!find('Trading Co') && !!find('Director 1'), 'Tabs built: '+book.sheets.map(s=>s.name).join(', '));
  ok(find('Trading Co').cells.A2==='Trading Co' && find('Director 1').cells.A2==='Director 1', 'Labels written into A2');
  ok(find('Entity template').visibility==='Hidden' && find('Person template').visibility==='Hidden', 'Templates hidden again');
  const order = book.sheets.map(s=>s.name); ok(order.indexOf('Trading Co') < order.indexOf('Liabilities'), 'New tabs placed before Liabilities');
  const hid = find('Servicing').hiddenRows; ok(!hid.has(6) && !hid.has(7) && hid.has(8) && hid.has(20), 'Unused Servicing party rows hidden');
  ok(!hid.has(59) && hid.has(71) && hid.has(83), 'IM blocks: entity shown, person + unused hidden');
  ok(/Built/.test(d.getElementById('svMsg').textContent), 'Build message: '+d.getElementById('svMsg').textContent.trim().slice(0,90));
  // paste P&L text
  w.handleText(`Profit and Loss
2025 2024
Trading Income
Sales 1,250,400 1,100,250
Total Trading Income 1,250,400 1,100,250
Cost of Sales
Opening Stock 40,000 35,000
Purchases 690,000 610,000
Closing Stock 45,000 40,000
Total Cost of Sales 685,000 605,000
Gross Profit 565,400 495,250
Other Income
Interest Income 1,200 900
Operating Expenses
Accounting Fees 8,500 8,000
Depreciation 25,000 22,000
Interest Expense 12,300 14,100
Rent 60,000 58,000
Wages and Salaries 310,000 278,600
Total Operating Expenses 415,800 380,700
Net Profit 150,800 115,450`);
  await tick();
  ok(vis('sendCard'), 'Send to workbook card shows after reading');
  ok(d.getElementById('svParty').selectedOptions[0].textContent==='Trading Co', 'Party defaults to Trading Co');
  ok(d.getElementById('svKind').value==='pl', 'Statement detected as P&L');
  const ys = [...d.querySelectorAll('#svYears select')].map(s=>s.selectedOptions[0].textContent); ok(ys.join()==='FY2025,FY2024', 'Year columns → '+ys.join(', '));
  ok(/Reconciles/.test(d.getElementById('svTotals').textContent), 'Preview reconciles to reported profit');
  d.getElementById('svWrite').click(); await tick(100);
  const t = find('Trading Co').cells;
  ok(t.D7===1250400 && t.C7===1100250 && t.D8===685000 && t.D10===1200 && t.D11===378500 && t.D13===25000 && t.D15===12300 && t.D17===150800, `P&L written: D7=${t.D7} D8=${t.D8} D11=${t.D11} D17=${t.D17}`);
  ok(t.B7===undefined, 'FY2023 column untouched');
  ok(/Wrote/.test(d.getElementById('svResult').textContent), 'Result: '+d.getElementById('svResult').textContent.trim());
  // add-to-existing on a second page
  d.getElementById('svAdd').checked = true;
  w.handleText(`Profit and Loss continued
2025 2024
Operating Expenses
Motor Vehicle 18,000 17,500`);
  await tick();
  d.getElementById('svWrite').click(); await tick(100);
  ok(find('Trading Co').cells.D11===396500, 'Second page added to opex: '+find('Trading Co').cells.D11);
  d.getElementById('svAdd').checked = false;
  // tax return → person
  w.handleText(`Individual tax return 2025
Salary or wages 145,000
Gross interest 1,250
Gross rent 32,000
Interest deductions 21,000
Taxable income 150,000
Tax on taxable income 36,000
Medicare levy 3,000`);
  await tick();
  ok(d.getElementById('svParty').selectedOptions[0].textContent==='Director 1', 'Tax return defaults to the individual');
  ok(d.getElementById('svKind').value==='tax', 'Statement detected as tax return');
  d.getElementById('svWrite').click(); await tick(100);
  const p = find('Director 1').cells;
  ok(p.D7===145000 && p.D9===1250 && p.D14===32000 && p.D15===21000 && p.D20===150000 && p.D21===39000, `Tax written: D7=${p.D7} D21=${p.D21}`);
  // manual override changes the target
  const sel = d.querySelector('#svMap select[data-map="1"]'); sel.value = '19'; sel.dispatchEvent(new w.Event('change',{bubbles:true})); await tick();
  ok(/Other income/.test(d.getElementById('svTotals').textContent), 'Manual re-match reflected in preview');
  console.log(fails ? `\n${fails} FAILED` : '\nALL PASSED');
  process.exit(fails?1:0);
})().catch(e=>{ console.error('ERROR', e); process.exit(1); });
