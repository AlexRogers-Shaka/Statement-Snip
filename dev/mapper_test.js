const P=require('./parser.js'), M=require('./mapper.js');
const toNum = s=>{ const r=P.toNumber(s,true); return r.v==null?null:r.v; };
function run(name, text, kind){
  const t=P.parseText(text); const k = kind||M.guessStatement(t.rows);
  const maps=M.autoMap(t.rows,k); const yc=M.guessYears(t.cols,['FY2023','FY2024','FY2025'],t.rows);
  console.log('\n==',name,'kind',k,'cols',t.cols,'years',yc);
  t.rows.forEach((r,i)=>console.log(String(maps[i].row??'-').padStart(3), (maps[i].neg?'NEG ':'')+r.cells[0], '|', r.cells.slice(1).join(' / '), maps[i].why?'  ('+maps[i].why+')':''));
  const agg=M.aggregate(t.rows,maps,yc,toNum); console.log(JSON.stringify(agg));
}
run('P&L2', `Trading Co Pty Ltd
Profit and Loss
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
Total Other Income 1,200 900
Operating Expenses
Accounting Fees 8,500 8,000
Bank Fees 1,200 1,100
Depreciation 25,000 22,000
Interest Expense 12,300 14,100
Motor Vehicle Expenses 18,000 17,500
Rent 60,000 58,000
Wages and Salaries 280,000 250,000
Superannuation 30,800 27,500
Total Operating Expenses 435,800 398,200
Net Profit 130,800 97,950`);
run('BS', `Balance Sheet
2025 2024
Assets
Current Assets
Cash at Bank 85,000 60,000
Trade Debtors 140,000 120,000
Stock on Hand 45,000 40,000
Total Current Assets 270,000 220,000
Non-current Assets
Plant and Equipment at cost 300,000 260,000
Less Accumulated Depreciation 110,000 85,000
Loan to Director 50,000 40,000
Total Non-current Assets 240,000 215,000
Total Assets 510,000 435,000
Liabilities
Current Liabilities
Trade Creditors 70,000 65,000
GST Payable 12,000 10,000
PAYG Withholding Payable 6,000 5,500
Hire Purchase - current 30,000 28,000
Total Current Liabilities 118,000 108,500
Non-current Liabilities
Hire Purchase 90,000 120,000
Bank Loan 150,000 160,000
Loan from Family Trust 40,000 0
Total Non-current Liabilities 280,000 280,000
Total Liabilities 398,000 388,500
Net Assets 112,000 46,500
Equity
Share Capital 100 100
Retained Earnings 111,900 46,400
Total Equity 112,000 46,500`);
run('Tax', `Individual tax return 2025
Salary or wages 145,000
Allowances, earnings, tips, director's fees 5,000
Gross interest 1,250
Dividends unfranked 800
Dividends franked 2,400
Franking credit 1,028
Distribution from trusts 40,000
Gross rent 32,000
Interest deductions 21,000
Capital works deductions 4,000
Other rental deductions 6,500
Total income or loss 202,950
Work related car expenses 1,200
Taxable income 190,250
Tax on taxable income 56,000
Medicare levy 3,805`);
run('P&L negative convention', `Profit and loss FY2025 FY2024
Revenue
Sales 500,000 450,000
Expenses
Wages (200,000) (180,000)
Depreciation (10,000) (9,000)
Interest (5,000) (4,000)
Net profit before tax 285,000 257,000`);
run('Tax2024', `Tax return for year ended 30 June 2024
Salary or wages 120,000
Tax on taxable income 30,000`);
