"""Pavilion Finance servicing workbook – tabs grow from the Structure tab.

Layout constants here are mirrored in the Statement Snip add-in (LAYOUT object).
"""
import sys, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

OUT, LOGO = sys.argv[1], sys.argv[2]

# ---------- brand ----------
NAVY, DEEP, CREAM = '173A53', '0D1B37', 'FFF9F2'
F, FS = 'Arial', 'Georgia'
BLUE = Font(name=F, size=10, color='0000FF')
BLACK = Font(name=F, size=10, color='1A1A1A')
GREEN = Font(name=F, size=10, color='008000')
BOLD = Font(name=F, size=10, bold=True, color='1A1A1A')
GREENB = Font(name=F, size=10, color='008000', bold=True)
TITLE = Font(name=FS, size=18, color=NAVY)
H2 = Font(name=FS, size=12, color=NAVY, bold=True)
SUB = Font(name=F, size=9, italic=True, color='6B7280')
HDR = Font(name=F, size=10, bold=True, color='FFFFFF')
SECF = Font(name=F, size=10, bold=True, color=NAVY)
HDRFILL = PatternFill('solid', fgColor=NAVY)
BANDFILL = PatternFill('solid', fgColor=DEEP)
SECFILL = PatternFill('solid', fgColor=CREAM)
INFILL = PatternFill('solid', fgColor='FFF2CC')
TOTFILL = PatternFill('solid', fgColor='EEF2F6')
thin = Side(style='thin', color='D0D5DD')
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
TOPB = Border(top=Side(style='thin', color=NAVY))
NAVYLINE = Border(bottom=Side(style='medium', color=NAVY))
CUR = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
MULT = '0.00"x";(0.00"x");"-"'
WRAP = Alignment(wrap_text=True, vertical='top')

N = 15                    # max parties
SF = 7                    # first Structure data row
SL = SF + N - 1           # 21
LF, LL = 7, 86            # liabilities rows
SYS_SHEETS = ['Cover', 'Guide', 'Assumptions', 'Structure', 'Entity template', 'Person template', 'Liabilities', 'Servicing']

wb = Workbook()

def inp(c, v=None, fmt=None):
    if v is not None: c.value = v
    c.font = BLUE; c.fill = INFILL; c.border = BOX
    if fmt: c.number_format = fmt
def fx(c, f, fmt=None, font=BLACK):
    c.value = f; c.font = font
    if fmt: c.number_format = fmt
def hdr(ws, row, vals, start=1, height=None):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=start+i, value=v); c.font = HDR; c.fill = HDRFILL
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    if height: ws.row_dimensions[row].height = height
def section(ws, row, text, ncol):
    for col in range(1, ncol+1):
        ws.cell(row=row, column=col).fill = SECFILL; ws.cell(row=row, column=col).border = Border(bottom=Side(style='thin', color=NAVY))
    ws.cell(row=row, column=1, value=text).font = SECF
def band(ws, ncol, title, sub):
    """Brand header: navy band row 1, title row 2, subtitle row 3."""
    for col in range(1, ncol+1):
        ws.cell(row=1, column=col).fill = HDRFILL
    ws.row_dimensions[1].height = 22
    c = ws.cell(row=1, column=1, value='PAVILION FINANCE'); c.font = Font(name=FS, size=11, color='FFFFFF', bold=True)
    c = ws.cell(row=1, column=ncol, value='Servicing assessment'); c.font = Font(name=F, size=9, color='FFFFFF'); c.alignment = Alignment(horizontal='right')
    if isinstance(title, str) and title.startswith('='): fx(ws['A2'], title, font=TITLE)
    else: ws['A2'] = title; ws['A2'].font = TITLE
    ws.row_dimensions[2].height = 28
    if sub is not None:
        if isinstance(sub, str) and sub.startswith('='): fx(ws['A3'], sub, font=Font(name=F, size=9, italic=True, color='008000'))
        else: ws['A3'] = sub; ws['A3'].font = SUB
def legend(ws, cell):
    ws[cell] = 'Cream-yellow cells with blue text are inputs · black = calculated · green = pulled from another tab'
    ws[cell].font = SUB
def finish(ws):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    ws.page_setup.orientation = 'portrait' if ws.title == 'Cover' else 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.oddFooter.left.text = 'Pavilion Finance – confidential'
    ws.oddFooter.right.text = 'Page &P of &N'

# ======================= Cover =======================
cv = wb.active; cv.title = 'Cover'
for col in range(1, 9): cv.column_dimensions[L(col)].width = 16
cv.column_dimensions['A'].width = 4
for r in range(1, 12):
    for col in range(1, 9): cv.cell(row=r, column=col).fill = PatternFill('solid', fgColor=CREAM)
img = XLImage(LOGO); img.width, img.height = 150, 133
cv.add_image(img, 'B2')
cv['E4'] = 'Servicing Assessment'; cv['E4'].font = Font(name=FS, size=24, color=NAVY)
cv['E6'] = 'Specialist Mortgage & Finance Brokers'; cv['E6'].font = Font(name=F, size=10, italic=True, color=NAVY)
fx(cv['E8'], '=IF(D14="","",D14)', font=Font(name=FS, size=14, color=NAVY))
for col in range(1, 9): cv.cell(row=11, column=col).border = NAVYLINE
cv['B13'] = 'Deal details'; cv['B13'].font = H2
details = [('Deal reference (anonymised)', 'Deal 001'), ('Prepared by', 'Pavilion Finance'), ('Date prepared', None),
           ('Proposed lender', None), ('Facility requested ($)', None), ('Purpose', None)]
for i, (lab, val) in enumerate(details):
    r = 14 + i
    cv.cell(row=r, column=2, value=lab).font = BLACK
    cv.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    c = cv.cell(row=r, column=4); inp(c, val)
    for cc in range(5, 8): cv.cell(row=r, column=cc).border = BOX; cv.cell(row=r, column=cc).fill = INFILL
cv['D16'].number_format = 'd mmmm yyyy'; cv['D18'].number_format = CUR
cv['B21'] = 'Snapshot'; cv['B21'].font = H2
snap_row = 22  # filled after Servicing built (needs row refs)
cv['B31'] = 'Client data stays on this computer. Use anonymised labels only – no names, TFNs, ABNs or addresses.'
cv['B31'].font = SUB
finish(cv)

# ======================= Guide =======================
gd = wb.create_sheet('Guide')
band(gd, 2, 'How to use this workbook', 'Set up once per deal, then snip statements straight into each party\'s tab with Statement Snip.')
steps = [
 ('1  Structure', 'List every entity and individual with an anonymised label, type, role and whether they are included in servicing. Paste a structure diagram below the table if you have one.'),
 ('2  Build tabs', 'In Excel: Home ▸ Statement Snip ▸ Servicing workbook ▸ Build tabs. One tab is created per party (entities get a financial statements tab, individuals and sole traders a tax return tab). Unused rows on the Servicing tab are hidden. Re-run it any time you add a party.'),
 ('   (no add-in?)', 'Right-click "Entity template" or "Person template" (Unhide first) ▸ Move or Copy ▸ Create a copy. Rename the copy to the party\'s "Tab name" from the Structure tab, and type the label into cell A2.'),
 ('3  Snip', 'Snip a P&L, balance sheet or tax return, paste it into Statement Snip and click Read. Under "Send to workbook", pick the party and year; each line is matched to the right row. Check the matching, then click Write to tab.'),
 ('4  Adjust', 'On each entity tab, add the servicing add-backs (one-off expenses, owner wages above market, non-recurring income), tax paid and committed outflows. Fill the reported net profit so the check line confirms nothing was missed.'),
 ('5  Liabilities', 'List every debt. Repayments at actual and assessment rates calculate automatically. Filter by party, type or status with the arrows on the header row.'),
 ('6  Servicing', 'The result: cash available, repayments, surplus, DSCR, interest cover, leverage, debt by type, and key financials by year for the information memorandum.'),
]
r = 5
for a, b in steps:
    gd.cell(row=r, column=1, value=a).font = BOLD
    c = gd.cell(row=r, column=2, value=b); c.font = BLACK; c.alignment = WRAP
    gd.row_dimensions[r].height = 44; r += 1
r += 1
gd.cell(row=r, column=1, value='Colour key').font = H2; r += 1
k = gd.cell(row=r, column=1, value='Input'); inp(k); gd.cell(row=r, column=2, value='Type here.').font = BLACK; r += 1
gd.cell(row=r, column=1, value='Calculated').font = BLACK; gd.cell(row=r, column=2, value='Formula – do not overwrite.').font = BLACK; r += 1
gd.cell(row=r, column=1, value='Linked').font = GREEN; gd.cell(row=r, column=2, value='Pulled from another tab.').font = BLACK; r += 2
gd.cell(row=r, column=1, value='Policy notes').font = H2; r += 1
notes = [
 'Lender policy differs. Assumptions-tab settings are placeholders – set them per lender before relying on results.',
 'Repayments: an actual monthly repayment from a statement is used if entered; otherwise it is calculated from the balance, rate and term.',
 'Assessed repayments: principal & interest over the remaining term on the limit, at the assessment rate. Credit cards use the % of limit setting; overdrafts and lines of credit are interest-only on the limit.',
 '"To be repaid" and "Refinance – being replaced" debts are excluded from servicing so the new facility is not double-counted.',
 'Individuals: remove distributions from entities already included in servicing (Person tab line 12) to avoid counting the same profit twice.',
 'Renaming a party after its tab is built: rename the tab to match the new "Tab name", or re-run Build tabs.',
]
for n in notes:
    c = gd.cell(row=r, column=2, value='• ' + n); c.font = BLACK; c.alignment = WRAP; gd.row_dimensions[r].height = 30; r += 1
gd.column_dimensions['A'].width = 18; gd.column_dimensions['B'].width = 110
finish(gd)

# ======================= Assumptions =======================
asn = wb.create_sheet('Assumptions')
band(asn, 11, 'Assumptions & lender policy settings', 'Placeholders, not any lender\'s policy. Set per lender.')
legend(asn, 'A4')
hdr(asn, 5, ['Setting', 'Value', 'Notes'])
rows = [
 (6, 'Financial year 1 (oldest)', 'FY2023', '@', 'Column headings on every party tab. Statement Snip matches snipped year columns to these.'),
 (7, 'Financial year 2', 'FY2024', '@', ''),
 (8, 'Financial year 3 (latest)', 'FY2025', '@', 'The "Latest year" basis uses this column.'),
 (9, 'Servicing basis', 'Latest year', None, 'Latest year, 2-year average or 3-year average (averages ignore blank years).'),
 (12, 'Assessment buffer (added to actual rate)', 0.03, PCT2, 'Placeholder.'),
 (13, 'Floor assessment rate', 0.0, PCT2, '0% = no floor. Placeholder.'),
 (14, 'Default remaining term (years)', 25, '0', 'Used when a debt has no remaining term.'),
 (15, 'Credit card assessed repayment (% of limit / month)', 0.038, PCT2, 'Placeholder.'),
 (16, 'Rental income shading (% of gross rent used)', 0.80, PCT, 'Placeholder.'),
 (18, 'DSCR benchmark (minimum)', 1.25, MULT, 'Placeholder.'),
 (19, 'Interest cover benchmark (minimum)', 2.0, MULT, 'Placeholder.'),
 (20, 'Business debt / EBITDA benchmark (maximum)', 3.0, MULT, 'Placeholder.'),
]
for rr, lab, val, fmt, note in rows:
    asn.cell(row=rr, column=1, value=lab).font = BLACK
    inp(asn.cell(row=rr, column=2), val, fmt)
    asn.cell(row=rr, column=3, value=note).font = SUB
asn.cell(row=10, column=1, value='Servicing basis number (calc)').font = BLACK
fx(asn['B10'], '=IFERROR(MATCH(B9,$E$6:$E$8,0),1)', '0')
section(asn, 11, 'Liability assessment', 3)
section(asn, 17, 'Benchmarks', 3)
lists = {
 'E': ('Servicing basis', ['Latest year', '2-year average', '3-year average']),
 'F': ('Party type', ['Company', 'Trust', 'Partnership', 'SMSF', 'Sole trader', 'Individual']),
 'G': ('Role', ['Borrower', 'Guarantor', 'Security provider', 'Related – not in deal']),
 'H': ('Debt type', ['Home loan', 'Investment property loan', 'Commercial property loan', 'Development / construction loan', 'Business term loan', 'Business overdraft', 'Line of credit', 'Equipment finance', 'Car loan', 'Credit card', 'Personal loan', 'ATO payment plan', 'Other']),
 'I': ('Status', ['Existing', 'Proposed', 'Refinance – being replaced', 'To be repaid']),
 'J': ('Repayment type', ['P&I', 'Interest only']),
 'K': ('Yes/No', ['Yes', 'No']),
}
for col, (name, vals) in lists.items():
    c = asn[f'{col}5']; c.value = name; c.font = HDR; c.fill = HDRFILL
    for i, v in enumerate(vals): asn[f'{col}{6+i}'] = v; asn[f'{col}{6+i}'].font = BLACK
asn['E4'] = 'Drop-down lists (add options here)'; asn['E4'].font = SUB
asn.column_dimensions['A'].width = 48; asn.column_dimensions['B'].width = 14; asn.column_dimensions['C'].width = 56; asn.column_dimensions['D'].width = 3
for col in 'EFGHIJK': asn.column_dimensions[col].width = 24
dv = DataValidation(type='list', formula1='=$E$6:$E$8', allow_blank=False); asn.add_data_validation(dv); dv.add('B9')
def LST(col, n): return f'=Assumptions!${col}$6:${col}${5+n}'
TYPES = lists['H'][1]
finish(asn)

# ======================= Structure =======================
st = wb.create_sheet('Structure')
band(st, 11, 'Structure', 'Use anonymised labels. Each labelled party gets its own tab (Statement Snip ▸ Build tabs).')
legend(st, 'A4')
hdr(st, 6, ['#', 'Label (anonymised)', 'Type', 'Owned / controlled by', 'Ownership %', 'Role in deal', 'Include in servicing?', 'Notes',
            'Tab name', 'Tab template', 'Tab status'], height=32)
for i in range(N):
    r = SF + i
    st.cell(row=r, column=1, value=i+1).font = BOLD
    for col in range(2, 9): inp(st.cell(row=r, column=col))
    st.cell(row=r, column=5).number_format = '0%'
    clean = f'B{r}'
    for ch in ['[', ']', ':', '*', '?', '/', '\\', "'"]:
        clean = f'SUBSTITUTE({clean},"{ch}","")'
    fx(st[f'I{r}'], f'=IF(B{r}="","",LEFT(TRIM({clean}),31))')
    fx(st[f'J{r}'], f'=IF(B{r}="","",IF(OR(C{r}="Individual",C{r}="Sole trader"),"Person","Entity"))')
    fx(st[f'K{r}'], f'=IF(I{r}="","",IF(ISREF(INDIRECT("\'"&I{r}&"\'!A1")),"Built","Not built yet"))')
st[f'B{SF}'] = 'Trading Co'; st[f'C{SF}'] = 'Company'; st[f'D{SF}'] = 'Family Trust'; st[f'E{SF}'] = 1.0
st[f'F{SF}'] = 'Borrower'; st[f'G{SF}'] = 'Yes'; st[f'H{SF}'] = 'Example – overwrite'
st[f'B{SF}'].comment = Comment('Example row showing the expected format. Overwrite with your own anonymised label.', 'Pavilion Finance')
for col, lst, n in [('C', 'F', 6), ('F', 'G', 4), ('G', 'K', 2)]:
    d = DataValidation(type='list', formula1=LST(lst, n), allow_blank=True); st.add_data_validation(d); d.add(f'{col}{SF}:{col}{SL}')
st.conditional_formatting.add(f'K{SF}:K{SL}', FormulaRule(formula=[f'K{SF}="Not built yet"'], font=Font(name=F, size=10, color='B45309', bold=True)))
st.conditional_formatting.add(f'K{SF}:K{SL}', FormulaRule(formula=[f'K{SF}="Built"'], font=Font(name=F, size=10, color='12805C', bold=True)))
for i, w in enumerate([5, 24, 14, 22, 11, 20, 12, 26, 24, 12, 14]): st.column_dimensions[L(i+1)].width = w
DIAG = SL + 3
st[f'A{DIAG}'] = 'Structure diagram – select the cell below and paste a screenshot (Ctrl+V)'; st[f'A{DIAG}'].font = H2
st.merge_cells(f'A{DIAG+1}:K{DIAG+24}')
for rr in range(DIAG+1, DIAG+25):
    for cc in range(1, 12): st.cell(row=rr, column=cc).border = BOX
st.freeze_panes = f'C{SF}'
finish(st)

# ======================= Templates =======================
YR = ["=Assumptions!$B$6", "=Assumptions!$B$7", "=Assumptions!$B$8"]
def basis(row): return f'=IFERROR(CHOOSE(Assumptions!$B$10,D{row},AVERAGE(C{row}:D{row}),AVERAGE(B{row}:D{row})),0)'
LOOK = lambda col: f'IFERROR(INDEX(Structure!${col}${SF}:${col}${SL},MATCH($A$2,Structure!$B${SF}:$B${SL},0)),"")'

def template_header(ws, kind):
    band(ws, 6, '(template – Statement Snip writes the party label here)',
         f'="Type: "&{LOOK("C")}&"   |   Role: "&{LOOK("F")}&"   |   Include in servicing: "&{LOOK("G")}')
    legend(ws, 'A4')
    hdr(ws, 5, ['Line item ($)' if kind == 'Entity' else 'Tax return item ($)', '', '', '', 'Servicing basis', 'Notes'])
    for j, f in enumerate(YR):
        c = ws.cell(row=5, column=2+j); fx(c, f, font=HDR); c.fill = HDRFILL; c.alignment = Alignment(horizontal='center')
    fx(ws['E5'], '=Assumptions!$B$9', font=HDR); ws['E5'].fill = HDRFILL; ws['E5'].alignment = Alignment(horizontal='center')

ENTITY_ROWS = [
 (6, 'Profit & loss', 'sec', None, ''),
 (7, 'Revenue', 'in', None, ''),
 (8, 'Cost of goods sold', 'in', None, ''),
 (9, 'Gross profit', 'calc', '={c}7-{c}8', ''),
 (10, 'Other income', 'in', None, ''),
 (11, 'Operating expenses (excl. depreciation, amortisation & interest)', 'in', None, 'Total expenses less D&A and interest'),
 (12, 'EBITDA', 'calc', '={c}9+{c}10-{c}11', ''),
 (13, 'Depreciation & amortisation', 'in', None, ''),
 (14, 'EBIT', 'calc', '={c}12-{c}13', ''),
 (15, 'Interest expense', 'in', None, ''),
 (16, 'Net profit before tax', 'calc', '={c}14-{c}15', ''),
 (17, 'Reported net profit before tax (per accounts)', 'in', None, 'From the accountant-prepared accounts'),
 (18, 'Check: calculated less reported (should be nil)', 'chk', '=IF({c}17="","",{c}16-{c}17)', 'Non-nil means a line is missing or mis-keyed'),
 (20, 'Servicing adjustments', 'sec', None, ''),
 (21, 'Add: non-recurring / one-off expenses', 'in', None, ''),
 (22, 'Add: owner wages & super above market rate', 'in', None, ''),
 (23, 'Less: non-recurring income', 'in', None, ''),
 (24, 'Adjusted EBITDA', 'calc', '={c}12+{c}21+{c}22-{c}23', ''),
 (25, 'Less: income tax paid', 'in', None, ''),
 (26, 'Less: other committed outflows (e.g. maintenance capex)', 'in', None, ''),
 (27, 'Cash available for debt servicing', 'calc', '={c}24-{c}25-{c}26', ''),
 (29, 'Balance sheet', 'sec', None, 'Servicing basis column = latest year'),
 (30, 'Cash', 'bin', None, ''),
 (31, 'Trade debtors', 'bin', None, ''),
 (32, 'Inventory / work in progress', 'bin', None, ''),
 (33, 'Other current assets', 'bin', None, ''),
 (34, 'Total current assets', 'bcalc', '=SUM({c}30:{c}33)', ''),
 (35, 'Property, plant & equipment', 'bin', None, ''),
 (36, 'Loans to related parties', 'bin', None, ''),
 (37, 'Other non-current assets', 'bin', None, ''),
 (38, 'Total assets', 'bcalc', '={c}34+SUM({c}35:{c}37)', ''),
 (39, 'Trade creditors', 'bin', None, ''),
 (40, 'Other current liabilities (excl. borrowings)', 'bin', None, ''),
 (41, 'Current borrowings', 'bin', None, ''),
 (42, 'Total current liabilities', 'bcalc', '=SUM({c}39:{c}41)', ''),
 (43, 'Non-current borrowings', 'bin', None, ''),
 (44, 'Loans from related parties', 'bin', None, ''),
 (45, 'Other non-current liabilities', 'bin', None, ''),
 (46, 'Total liabilities', 'bcalc', '={c}42+SUM({c}43:{c}45)', ''),
 (47, 'Net assets (equity)', 'bcalc', '={c}38-{c}46', ''),
 (49, 'Ratios', 'sec', None, ''),
 (50, 'Revenue growth', 'growth', None, ''),
 (51, 'Gross margin', 'ratio', '=IFERROR({c}9/{c}7,"")', ''),
 (52, 'EBITDA margin (adjusted)', 'ratio', '=IFERROR({c}24/{c}7,"")', ''),
 (53, 'Net profit margin (before tax)', 'ratio', '=IFERROR({c}16/{c}7,"")', ''),
 (54, 'Current ratio', 'mult', '=IFERROR({c}34/{c}42,"")', ''),
 (55, 'Working capital', 'bcalc', '={c}34-{c}42', ''),
 (56, 'Total borrowings', 'bcalc', '={c}41+{c}43', ''),
 (57, 'Debt / adjusted EBITDA', 'mult', '=IFERROR({c}56/{c}24,"")', ''),
 (58, 'Debt / equity (gearing)', 'mult', '=IFERROR({c}56/{c}47,"")', ''),
 (59, 'Interest cover (adjusted EBITDA / interest)', 'mult', '=IFERROR({c}24/{c}15,"")', ''),
]
BOLDROWS = {9, 12, 14, 16, 24, 27, 34, 38, 42, 46, 47}
et = wb.create_sheet('Entity template')
template_header(et, 'Entity')
for row, lab, kind, tmpl, note in ENTITY_ROWS:
    if kind == 'sec':
        section(et, row, lab, 6)
        if note: et[f'F{row}'] = note; et[f'F{row}'].font = SUB
        continue
    et.cell(row=row, column=1, value=lab).font = BOLD if row in BOLDROWS else BLACK
    for col in 'BCD':
        c = et[f'{col}{row}']
        if kind in ('in', 'bin'): inp(c, None, CUR)
        elif kind == 'growth':
            prev = {'B': None, 'C': 'B', 'D': 'C'}[col]
            fx(c, None if prev is None else f'=IFERROR(IF(OR({prev}7="",{prev}7=0,{col}7=""),"",{col}7/{prev}7-1),"")', PCT)
        else:
            fmt = PCT if kind == 'ratio' else MULT if kind == 'mult' else CUR
            fx(c, tmpl.format(c=col), fmt, BOLD if row in BOLDROWS else BLACK)
            if kind == 'calc': c.border = TOPB
    e = et[f'E{row}']
    if kind == 'in': fx(e, basis(row), CUR)
    elif kind in ('calc', 'chk'):
        fx(e, tmpl.format(c='E'), CUR, BOLD if row in BOLDROWS else BLACK)
        if kind == 'calc': e.border = TOPB
    elif kind == 'bin': fx(e, f'=D{row}', CUR)
    elif kind == 'bcalc': fx(e, tmpl.format(c='E'), CUR, BOLD if row in BOLDROWS else BLACK)
    elif kind == 'ratio': fx(e, tmpl.format(c='E'), PCT)
    elif kind == 'mult': fx(e, tmpl.format(c='E'), MULT)
    if note: et[f'F{row}'] = note; et[f'F{row}'].font = SUB
et.conditional_formatting.add('B18:E18', FormulaRule(formula=['AND(B18<>"",ABS(B18)>=1)'], font=Font(name=F, size=10, color='C2410C', bold=True), fill=PatternFill('solid', fgColor='FDECEC')))
et.column_dimensions['A'].width = 58
for col in 'BCDE': et.column_dimensions[col].width = 15
et.column_dimensions['F'].width = 42
et.freeze_panes = 'B6'
finish(et); et.sheet_state = 'hidden'

PERSON_ROWS = [
 (6, 'Income per tax return', 'sec', None, ''),
 (7, 'Item 1 – Salary or wages', 'in', None, ''),
 (8, 'Item 2 – Allowances, earnings, tips, director\'s fees', 'in', None, ''),
 (9, 'Item 10 – Gross interest', 'in', None, ''),
 (10, 'Item 11 – Dividends (franked + unfranked)', 'in', None, 'Exclude franking credits'),
 (11, 'Item 13 – Partnership & trust distributions', 'in', None, ''),
 (12, 'Less: distributions from group entities already in servicing', 'in', None, 'Avoids double-counting entity profits'),
 (13, 'Item 15 / P8 – Net business income (sole trader)', 'in', None, ''),
 (14, 'Item 21 – Gross rent', 'in', None, ''),
 (15, 'Item 21 – Rental interest deductions', 'in', None, ''),
 (16, 'Item 21 – Capital works deductions', 'in', None, ''),
 (17, 'Item 21 – Other rental deductions', 'in', None, ''),
 (18, 'Net rent per return', 'calc', '={c}14-{c}15-{c}16-{c}17', ''),
 (19, 'Other income', 'in', None, ''),
 (20, 'Taxable income (per return / NOA)', 'in', None, 'Cross-check only'),
 (21, 'Tax payable incl. Medicare levy (per NOA)', 'in', None, ''),
 (23, 'Servicing', 'sec', None, ''),
 (24, 'Rental income used (gross rent × shading)', 'calc', '={c}14*Assumptions!$B$16', 'Rental debt is assessed on the Liabilities tab'),
 (25, 'Gross income used for servicing', 'calc', '={c}7+{c}8+{c}9+{c}10+{c}11-{c}12+{c}13+{c}24+{c}19', ''),
 (26, 'Less: tax payable', 'calc', '={c}21', ''),
 (27, 'Net income for servicing', 'calc', '={c}25-{c}26', ''),
]
pt = wb.create_sheet('Person template')
template_header(pt, 'Person')
for row, lab, kind, tmpl, note in PERSON_ROWS:
    if kind == 'sec': section(pt, row, lab, 6); continue
    bold = row in (18, 25, 27)
    pt.cell(row=row, column=1, value=lab).font = BOLD if bold else BLACK
    for col in 'BCD':
        c = pt[f'{col}{row}']
        if kind == 'in': inp(c, None, CUR)
        else:
            fx(c, tmpl.format(c=col), CUR, BOLD if bold else BLACK)
            if row in (18, 27): c.border = TOPB
    e = pt[f'E{row}']
    if kind == 'in': fx(e, basis(row), CUR)
    else:
        fx(e, tmpl.format(c='E'), CUR, BOLD if bold else BLACK)
        if row in (18, 27): e.border = TOPB
    if note: pt[f'F{row}'] = note; pt[f'F{row}'].font = SUB
pt.column_dimensions['A'].width = 58
for col in 'BCDE': pt.column_dimensions[col].width = 15
pt.column_dimensions['F'].width = 42
pt.freeze_panes = 'B6'
finish(pt); pt.sheet_state = 'hidden'

# ======================= Liabilities =======================
li = wb.create_sheet('Liabilities')
band(li, 20, 'Liabilities', 'All debts across the group. Filter with the arrows on the header row; totals show visible rows only.')
legend(li, 'A4')
heads = ['#', 'Entity / person', 'Debt type', 'Lender', 'Purpose / security', 'Status', 'Limit ($)', 'Balance owing ($)',
         'Interest rate (actual)', 'Repayment type', 'Remaining term (yrs)', 'Actual monthly repayment per statement ($)',
         'Monthly repayment – actual ($)', 'Assessment rate override', 'Assessment rate', 'Monthly repayment – assessed ($)',
         'Annual repayment – actual ($)', 'Annual repayment – assessed ($)', 'Include in servicing?', 'Comments']
hdr(li, 6, heads, height=56)
for rr in range(LF, LL+1):
    li.cell(row=rr, column=1, value=rr-LF+1).font = BLACK
    for col in list(range(2, 13)) + [14, 20]: inp(li.cell(row=rr, column=col))
    li[f'G{rr}'].number_format = CUR; li[f'H{rr}'].number_format = CUR; li[f'I{rr}'].number_format = PCT2
    li[f'K{rr}'].number_format = '0.0'; li[f'L{rr}'].number_format = CUR; li[f'N{rr}'].number_format = PCT2
    base = f'IF(H{rr}<>"",H{rr},G{rr})'
    term = f'IF(K{rr}<>"",K{rr},Assumptions!$B$14)'
    fx(li[f'M{rr}'], f'=IF(AND(G{rr}="",H{rr}=""),"",IF(L{rr}<>"",L{rr},IF(C{rr}="Credit card",MAX(G{rr},H{rr})*Assumptions!$B$15,'
                     f'IF(OR(C{rr}="Business overdraft",C{rr}="Line of credit",J{rr}="Interest only"),{base}*I{rr}/12,'
                     f'IFERROR(PMT(I{rr}/12,{term}*12,-{base}),0)))))', CUR)
    fx(li[f'O{rr}'], f'=IF(AND(G{rr}="",H{rr}=""),"",IF(N{rr}<>"",N{rr},MAX(I{rr}+Assumptions!$B$12,Assumptions!$B$13)))', PCT2)
    lim = f'MAX(G{rr},H{rr})'
    fx(li[f'P{rr}'], f'=IF(AND(G{rr}="",H{rr}=""),"",IF(C{rr}="Credit card",{lim}*Assumptions!$B$15,'
                     f'IF(OR(C{rr}="Business overdraft",C{rr}="Line of credit"),{lim}*O{rr}/12,'
                     f'IFERROR(PMT(O{rr}/12,{term}*12,-{lim}),0))))', CUR)
    fx(li[f'Q{rr}'], f'=IF(M{rr}="","",M{rr}*12)', CUR)
    fx(li[f'R{rr}'], f'=IF(P{rr}="","",P{rr}*12)', CUR)
    fx(li[f'S{rr}'], f'=IF(B{rr}="","",IF(OR(F{rr}="To be repaid",F{rr}="Refinance – being replaced"),"No","Yes"))')
ex = {'B': 'Trading Co', 'C': 'Equipment finance', 'D': 'Lender A', 'E': 'Excavator', 'F': 'Existing', 'G': 150000, 'H': 120000,
      'I': 0.075, 'J': 'P&I', 'K': 4, 'T': 'Example – overwrite or delete'}
for col, v in ex.items(): li[f'{col}{LF}'].value = v
li[f'B{LF}'].comment = Comment('Example row showing the expected format. Overwrite or delete it.', 'Pavilion Finance')
for col, f1 in [('B', f'=Structure!$B${SF}:$B${SL}'), ('C', LST('H', len(TYPES))), ('F', LST('I', 4)), ('J', LST('J', 2))]:
    d = DataValidation(type='list', formula1=f1, allow_blank=True); li.add_data_validation(d); d.add(f'{col}{LF}:{col}{LL}')
li['F5'] = 'Totals (visible rows):'; li['F5'].font = BOLD; li['F5'].alignment = Alignment(horizontal='right')
for col in 'GHMPQR':
    c = li[f'{col}5']; fx(c, f'=SUBTOTAL(109,{col}{LF}:{col}{LL})', CUR, BOLD); c.fill = TOTFILL; c.border = BOX
li.auto_filter.ref = f'A6:T{LL}'
for i, x in enumerate([5, 18, 24, 14, 22, 22, 14, 14, 11, 12, 11, 16, 15, 12, 11, 15, 15, 15, 11, 36]):
    li.column_dimensions[L(i+1)].width = x
li.freeze_panes = f'C{LF}'
finish(li)

# ======================= Servicing =======================
sv = wb.create_sheet('Servicing')
band(sv, 14, 'Servicing summary', None)
fx(sv['A3'], '="Basis: "&Assumptions!$B$9&"   |   Buffer: "&TEXT(Assumptions!$B$12,"0.00%")&"   |   Floor: "&TEXT(Assumptions!$B$13,"0.00%")&"   |   Deal: "&Cover!$D$14', font=GREEN)
hdr(sv, 5, ['#', 'Party', 'Tab type', 'Role', 'Include?', 'Revenue ($)', 'Adjusted EBITDA ($)', 'Interest expense ($)', 'Cash available / net income ($)',
            'Debt limits ($)', 'Annual repayments – actual ($)', 'Annual repayments – assessed ($)', 'DSCR – actual', 'DSCR – assessed'], height=44)
P0 = 6; P1 = P0 + N - 1       # 6..20
LR = f'Liabilities!$B${LF}:$B${LL}'; LS = f'Liabilities!$S${LF}:$S${LL}'
def ind(r, cell, dflt='0'):  # value from the party's tab, default if not built
    ref = f'"\'"&Structure!$I${r}&"\'!{cell}"'
    return f'IF(ISREF(INDIRECT({ref})),N(INDIRECT({ref})),{dflt})'
for i in range(N):
    rr = P0 + i; s = SF + i
    sv.cell(row=rr, column=1, value=i+1).font = BLACK
    fx(sv[f'B{rr}'], f'=IF(Structure!$B${s}="","",Structure!$B${s})', font=GREEN)
    fx(sv[f'C{rr}'], f'=Structure!$J${s}', font=GREEN)
    fx(sv[f'D{rr}'], f'=IF(Structure!$F${s}="","",Structure!$F${s})', font=GREEN)
    fx(sv[f'E{rr}'], f'=IF(B{rr}="","",IF(Structure!$G${s}="","No",Structure!$G${s}))', font=GREEN)
    fx(sv[f'F{rr}'], f'=IF(C{rr}="Entity",{ind(s,"$E$7")},"")', CUR, GREEN)
    fx(sv[f'G{rr}'], f'=IF(C{rr}="Entity",{ind(s,"$E$24")},"")', CUR, GREEN)
    fx(sv[f'H{rr}'], f'=IF(C{rr}="Entity",{ind(s,"$E$15")},"")', CUR, GREEN)
    fx(sv[f'I{rr}'], f'=IF(B{rr}="","",{ind(s,"$E$27")})', CUR, GREEN)
    fx(sv[f'J{rr}'], f'=IF(B{rr}="","",SUMIFS(Liabilities!$G${LF}:$G${LL},{LR},B{rr},{LS},"Yes"))', CUR)
    fx(sv[f'K{rr}'], f'=IF(B{rr}="","",SUMIFS(Liabilities!$Q${LF}:$Q${LL},{LR},B{rr},{LS},"Yes"))', CUR)
    fx(sv[f'L{rr}'], f'=IF(B{rr}="","",SUMIFS(Liabilities!$R${LF}:$R${LL},{LR},B{rr},{LS},"Yes"))', CUR)
    fx(sv[f'M{rr}'], f'=IFERROR(I{rr}/K{rr},"")', MULT)
    fx(sv[f'N{rr}'], f'=IFERROR(I{rr}/L{rr},"")', MULT)
def incl(col, typ=None):
    t = f',$C${P0}:$C${P1},"{typ}"' if typ else ''
    return f'SUMIFS({col}{P0}:{col}{P1},$E${P0}:$E${P1},"Yes"{t})'
rr = P1 + 2
section(sv, rr, 'Group servicing (parties with Include = Yes)', 6); rr += 1
def line(label, f, fmt=CUR, bold=False, inputcell=False, note=None):
    global rr
    sv.cell(row=rr, column=2, value=label).font = BOLD if bold else BLACK
    c = sv[f'F{rr}']
    if inputcell: inp(c, f, fmt)
    else: fx(c, f, fmt, BOLD if bold else BLACK)
    if bold: c.border = TOPB
    if note: sv[f'G{rr}'] = note; sv[f'G{rr}'].font = SUB
    rr += 1; return rr - 1
r_ent = line('Business cash available for debt servicing', f'={incl("I","Entity")}')
r_per = line('Personal net income for servicing', f'={incl("I","Person")}')
r_liv = line('Less: household living expenses (annual)', 0, inputcell=True, note='Declared expenses or the lender\'s benchmark, whichever is higher')
r_tot = line('Total cash available for debt servicing', f'=F{r_ent}+F{r_per}-F{r_liv}', bold=True)
rr += 1
r_ra = line('Annual repayments – actual', f'={incl("K")}')
r_rs = line('Annual repayments – assessed', f'={incl("L")}')
r_sa = line('Surplus / (shortfall) – actual', f'=F{r_tot}-F{r_ra}', bold=True)
r_ss = line('Surplus / (shortfall) – assessed', f'=F{r_tot}-F{r_rs}', bold=True)
rr += 1
section(sv, rr, 'Key ratios', 6); rr += 1
for col, t in zip('FGH', ['Result', 'Benchmark', 'Outcome']):
    sv[f'{col}{rr}'] = t; sv[f'{col}{rr}'].font = BOLD
rr += 1
RATIO_ROWS = {}
def ratio(key, label, f, bench, higher_good=True):
    global rr
    sv.cell(row=rr, column=2, value=label).font = BLACK
    fx(sv[f'F{rr}'], f, MULT, BOLD); fx(sv[f'G{rr}'], bench, MULT, GREEN)
    op = '>=' if higher_good else '<='
    fx(sv[f'H{rr}'], f'=IF(F{rr}="","",IF(F{rr}{op}G{rr},"Meets","Outside"))', None, BOLD)
    RATIO_ROWS[key] = rr; rr += 1
ratio('dscr_a', 'DSCR – actual rates', f'=IFERROR(F{r_tot}/F{r_ra},"")', '=Assumptions!$B$18')
ratio('dscr_s', 'DSCR – assessment rates', f'=IFERROR(F{r_tot}/F{r_rs},"")', '=Assumptions!$B$18')
ratio('icr', 'Interest cover (adjusted EBITDA ÷ interest, included entities)', f'=IFERROR({incl("G","Entity")}/{incl("H","Entity")},"")', '=Assumptions!$B$19')
ratio('lev', 'Business debt limits ÷ adjusted EBITDA (included entities)', f'=IFERROR({incl("J","Entity")}/{incl("G","Entity")},"")', '=Assumptions!$B$20', higher_good=False)
sv.cell(row=rr, column=2, value='Total debt limits (included parties)').font = BLACK
fx(sv[f'F{rr}'], f'={incl("J")}', CUR, BOLD); r_debt = rr; rr += 1
for key in RATIO_ROWS:
    r = RATIO_ROWS[key]
    sv.conditional_formatting.add(f'H{r}', FormulaRule(formula=[f'H{r}="Meets"'], font=Font(name=F, size=10, color='12805C', bold=True)))
    sv.conditional_formatting.add(f'H{r}', FormulaRule(formula=[f'H{r}="Outside"'], font=Font(name=F, size=10, color='C2410C', bold=True)))
rr += 1
section(sv, rr, 'Debt by type (included in servicing)', 9); rr += 1
hdr(sv, rr, ['', 'Debt type', '', '', '', 'Limits ($)', 'Annual – actual ($)', 'Annual – assessed ($)', '% of limits']); rr += 1
t0 = rr
for t in TYPES:
    sv.cell(row=rr, column=2, value=t).font = BLACK
    for col, src in [('F', 'G'), ('G', 'Q'), ('H', 'R')]:
        fx(sv[f'{col}{rr}'], f'=SUMIFS(Liabilities!${src}${LF}:${src}${LL},Liabilities!$C${LF}:$C${LL},$B{rr},{LS},"Yes")', CUR)
    rr += 1
t1 = rr - 1
for r_ in range(t0, t1+1): fx(sv[f'I{r_}'], f'=IFERROR(F{r_}/F${t1+1},"")', PCT)
sv.cell(row=rr, column=2, value='Total').font = BOLD
for col in 'FGH': fx(sv[f'{col}{rr}'], f'=SUM({col}{t0}:{col}{t1})', CUR, BOLD); sv[f'{col}{rr}'].border = TOPB
fx(sv[f'I{rr}'], f'=IFERROR(F{rr}/F{rr},"")', PCT, BOLD)
rr += 2
section(sv, rr, 'Key financials by year (for the IM)', 9); rr += 1
IM_FIRST = rr
IM_ITEMS = [('Revenue', 7, CUR), ('Revenue growth', 50, PCT), ('Gross margin', 51, PCT), ('Adjusted EBITDA', 24, CUR), ('EBITDA margin', 52, PCT),
            ('Net profit before tax', 16, CUR), ('Net assets', 47, CUR), ('Current ratio', 54, MULT), ('Debt / adjusted EBITDA', 57, MULT),
            ('Interest cover', 59, MULT)]
IM_BLOCK = len(IM_ITEMS) + 2
for i in range(N):
    s = SF + i; top = IM_FIRST + i * IM_BLOCK
    fx(sv[f'B{top}'], f'=IF(Structure!$J${s}="Entity",Structure!$B${s},"")', font=GREENB)
    for j, col in enumerate('CDE'):
        fx(sv[f'{col}{top}'], f'=IF($B{top}="","",{YR[j][1:]})', font=BOLD); sv[f'{col}{top}'].alignment = Alignment(horizontal='right')
    fx(sv[f'F{top}'], f'=IF($B{top}="","",Assumptions!$B$9)', font=BOLD); sv[f'F{top}'].alignment = Alignment(horizontal='right')
    for k, (lab, srow, fmt) in enumerate(IM_ITEMS):
        r = top + 1 + k
        fx(sv[f'B{r}'], f'=IF($B${top}="","","   {lab}")')
        for col, src in zip('CDEF', 'BCDE'):
            fx(sv[f'{col}{r}'], f'=IF($B${top}="","",IF(ISREF(INDIRECT("\'"&Structure!$I${s}&"\'!${src}${srow}")),INDIRECT("\'"&Structure!$I${s}&"\'!${src}${srow}"),""))', fmt, GREEN)
sv.column_dimensions['A'].width = 4; sv.column_dimensions['B'].width = 44
for col, wd in zip('CDEFGHIJKLMN', [12, 12, 12, 16, 16, 16, 18, 15, 16, 16, 11, 11]): sv.column_dimensions[col].width = wd
sv.freeze_panes = 'C6'
finish(sv)

# ---------- cover snapshot ----------
snap = [('Total cash available for debt servicing', f'=Servicing!F{r_tot}', CUR),
        ('Annual repayments – assessed', f'=Servicing!F{r_rs}', CUR),
        ('Surplus / (shortfall) – assessed', f'=Servicing!F{r_ss}', CUR),
        ('DSCR – actual', f'=Servicing!F{RATIO_ROWS["dscr_a"]}', MULT),
        ('DSCR – assessed', f'=Servicing!F{RATIO_ROWS["dscr_s"]}', MULT),
        ('Interest cover', f'=Servicing!F{RATIO_ROWS["icr"]}', MULT),
        ('Total debt limits', f'=Servicing!F{r_debt}', CUR)]
for i, (lab, f, fmt) in enumerate(snap):
    r = snap_row + i
    cv.cell(row=r, column=2, value=lab).font = BLACK
    c = cv.cell(row=r, column=5); fx(c, f, fmt, Font(name=F, size=11, bold=True, color=NAVY)); c.alignment = Alignment(horizontal='right')
    cv.cell(row=r, column=6, value='').font = BLACK
    if lab.startswith('DSCR'):
        fx(cv.cell(row=r, column=6), f'=IF(E{r}="","",IF(E{r}>=Assumptions!$B$18,"Meets benchmark","Below benchmark"))', None, Font(name=F, size=9, italic=True, color=NAVY))
    for cc in range(2, 7): cv.cell(row=r, column=cc).border = Border(bottom=thin)

wb._sheets = [wb[n] for n in SYS_SHEETS]
wb.active = 0
wb.save(OUT)
layout = {'structFirst': SF, 'n': N, 'servFirst': P0, 'imFirst': IM_FIRST, 'imBlock': IM_BLOCK,
          'liabFirst': LF, 'liabLast': LL, 'systemSheets': SYS_SHEETS}
json.dump(layout, open(OUT + '.layout.json', 'w'))
print('saved', OUT, layout)
