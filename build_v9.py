"""
Rebuild Lending_Business_System_V8 -> V9.

Fixes the defects found in the V8 audit, adds dropdown validation, fills the
dead columns, and applies consistent professional formatting.
Run:  python3 build_v9.py
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import TableColumn
from openpyxl.utils import range_boundaries
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.datavalidation import DataValidationList

SRC = "Lending_Business_System_V8 recreated.xlsx"
DST = "Lending_Business_System_V9.xlsx"
R0, R1 = 9, 508                      # data rows
ROWS = range(R0, R1 + 1)

# ---------------------------------------------------------------- palette ---
NAVY   = "1F3864"; MIDBLUE = "2E5A8A"; PALEBLUE = "D9E2F3"
GREY   = "F2F2F2"; LINE    = "BFBFBF"
INPUT_FONT = "0070C0"; LINK_FONT = "375623"; LINK_FILL = "E2EFDA"
RED_F, RED_T   = "FFC7CE", "9C0006"
YEL_F, YEL_T   = "FFEB9C", "9C6500"
GRN_F, GRN_T   = "C6EFCE", "006100"

MONEY = '#,##0.00;[Red]\\(#,##0.00\\);"–"'
INT   = '#,##0;[Red]\\(#,##0\\);"–"'
PCT   = '0.0%;[Red]\\(0.0%\\);"–"'
DATE  = 'dd-mmm-yyyy'

thin  = Side(style="thin", color=LINE)
BOX   = Border(left=thin, right=thin, top=thin, bottom=thin)
BOX_B = Border(left=thin, right=thin, top=thin, bottom=Side(style="medium", color=NAVY))

F_TITLE = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
F_SUB   = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
F_BAND  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
F_HDR   = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_BODY  = Font(name="Calibri", size=10)
F_INPUT = Font(name="Calibri", size=10, color=INPUT_FONT, bold=True)
F_CALC  = Font(name="Calibri", size=10, color="404040")
F_LINK  = Font(name="Calibri", size=10, color=LINK_FONT)
F_KPI_L = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_KPI_V = Font(name="Calibri", size=20, bold=True, color=NAVY)
F_SEC   = Font(name="Calibri", size=11, bold=True, color=NAVY)

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_BAND  = PatternFill("solid", fgColor=MIDBLUE)
FILL_HDR   = PatternFill("solid", fgColor=NAVY)
FILL_CALC  = PatternFill("solid", fgColor=GREY)
FILL_LINK  = PatternFill("solid", fgColor=LINK_FILL)
FILL_KPI   = PatternFill("solid", fgColor=PALEBLUE)
FILL_WHITE = PatternFill("solid", fgColor="FFFFFF")

CTR  = Alignment(horizontal="center", vertical="center")
CTRW = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT  = Alignment(horizontal="left", vertical="center")
RGT  = Alignment(horizontal="right", vertical="center")


# ---------------------------------------------------------------- helpers ---
def banner(ws, title, subtitle, ncols, height=30):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(1, 1, title); c.font = F_TITLE; c.fill = FILL_TITLE; c.alignment = LFT
    c2 = ws.cell(2, 1, subtitle); c2.font = F_SUB; c2.fill = FILL_TITLE; c2.alignment = LFT
    for col in range(1, ncols + 1):
        ws.cell(1, col).fill = FILL_TITLE
        ws.cell(2, col).fill = FILL_TITLE
    ws.row_dimensions[1].height = height
    ws.row_dimensions[2].height = 16


def section(ws, text, subtitle, last_col=14):
    """Rows 4/5/6 = section heading, instruction line, summary-strip labels."""
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=last_col)
    c = ws.cell(4, 1, text); c.font = F_BAND; c.fill = FILL_BAND; c.alignment = LFT
    for col in range(1, last_col + 1):
        ws.cell(4, col).fill = FILL_BAND
    ws.row_dimensions[4].height = 20
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=last_col)
    c = ws.cell(5, 1, subtitle)
    c.font = Font(name="Calibri", size=9, italic=True, color="595959"); c.alignment = LFT
    ws.row_dimensions[5].height = 15


def summary_pairs(ws, pairs, row=7):
    """pairs = [(label, formula, numfmt), ...] laid out two columns each."""
    col = 1
    for label, formula, fmt in pairs:
        lc = ws.cell(row, col, label)
        lc.font = F_HDR; lc.fill = FILL_HDR; lc.alignment = LFT
        vc = ws.cell(row, col + 1, formula)
        vc.font = Font(name="Calibri", size=12, bold=True, color=NAVY)
        vc.fill = FILL_KPI; vc.alignment = RGT; vc.number_format = fmt
        vc.border = BOX
        ws.column_dimensions[get_column_letter(col)].width = max(
            ws.column_dimensions[get_column_letter(col)].width or 0, len(label) + 2)
        col += 2
    ws.row_dimensions[row].height = 22


def header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row, start_col + i, h)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTRW; c.border = BOX_B
    ws.row_dimensions[row].height = 34


def style_data(ws, cols, kinds, widths, numfmts):
    """kinds: dict col_index -> 'input'|'calc'|'link'|'flag'."""
    for r in ROWS:
        for ci in cols:
            c = ws.cell(r, ci)
            k = kinds.get(ci, "input")
            c.border = BOX
            c.font = {"input": F_INPUT, "calc": F_CALC, "link": F_LINK, "flag": F_BODY}[k]
            if k in ("calc", "flag"):
                c.fill = FILL_CALC
            elif k == "link":
                c.fill = FILL_LINK
            if ci in numfmts:
                c.number_format = numfmts[ci]
            if k == "flag":
                c.alignment = CTR
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def put_formulas(ws, col, template):
    for r in ROWS:
        ws.cell(r, col, template.format(r=r))


def flag_cf(ws, rng, text, fill, txt):
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{rng.split(":")[0]}="{text}"'],
        fill=PatternFill("solid", bgColor=fill), font=Font(color=txt, bold=True),
        stopIfTrue=False))


def money_cf(ws, rng):
    first = rng.split(":")[0]
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="greaterThan", formula=["0"],
        fill=PatternFill("solid", bgColor=PALEBLUE)))


# =============================================================== load ======
wb = openpyxl.load_workbook(SRC)

# Reset the sheets we rebuild. The two printable paper forms are left exactly
# as they are - they are standalone printouts, not part of the calculation.
KEEP = {"COLLECTION SHEET", "Daily Cash Report Format"}
for ws in wb.worksheets:
    if ws.title in KEEP:
        continue
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws.conditional_formatting = ConditionalFormattingList()
    ws.data_validations = DataValidationList()
    for row in ws.iter_rows():
        for c in row:
            c.value = None
    # V8 hid several columns; some now carry live content or user input.
    for dim in ws.column_dimensions.values():
        dim.hidden = False

# ---------------------------------------------------------------- Lists ----
ls = wb["Lists"]
for row in ls.iter_rows():
    for c in row:
        c.value = None
        c.fill = PatternFill(); c.font = Font(); c.border = Border()

# Row 1-2 are the banner, so the lists cannot start there - a merged banner
# cell would swallow the header row and the first item of every list.
HDR_ROW, FIRST_ITEM = 4, 5

lists_cols = {
    1: ("BorrowerStatus",    "Borrower Status", ["ACTIVE", "INACTIVE", "BLOCKED"]),
    2: ("PaymentMethod",     "Payment Method",  ["Cash", "GCash", "Bank Transfer", "Cheque"]),
    3: ("ExpenseCategory",   "Expense Category",["Transportation", "Supplies", "Utilities",
                                                "Salaries & Wages", "Rent", "Maintenance",
                                                "Food/Meals", "Marketing", "Taxes & Permits",
                                                "PCF", "Other"]),
    4: ("LoanStatusList",    "Loan Status",     ["ACTIVE", "PAID", "PAST DUE"]),
    6: ("CollectorNames",    "Payment Collector",["Collector 1", "Collector 2", "Collector 3", "Owner"]),
    7: ("OfficerNames",      "Loan Officer",    ["Officer 1", "Officer 2", "Officer 3", "Owner"]),
    9: ("InterestBasis",     "Interest Basis",  ["Daily", "Weekly", "Monthly", "Yearly"]),
   10: ("RepayFrequency",    "Repayment Frequency", ["Daily", "Weekly", "Monthly"]),
   11: ("DisburseChannel",   "Disbursement Channel", ["Cash", "GCash", "Bank Transfer"]),
   12: ("ApproverNames",     "Approved By",     ["Owner", "Manager", "Officer 1"]),
}
for ci, (_key, title, items) in lists_cols.items():
    c = ls.cell(HDR_ROW, ci, title)
    c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTRW; c.border = BOX_B
    ls.column_dimensions[get_column_letter(ci)].width = 22
    for i, v in enumerate(items):
        rr = FIRST_ITEM + i
        cc = ls.cell(rr, ci, v)
        cc.font = F_BODY; cc.border = BOX; cc.alignment = LFT
        if i % 2:
            cc.fill = PatternFill("solid", fgColor=GREY)
ls.row_dimensions[HDR_ROW].height = 30
ls.column_dimensions["E"].width = 3
ls.column_dimensions["H"].width = 3

notes = [
    "V9 CONTROL NOTES",
    "Never delete or rename an ID once it has been used. Add a new row for a new transaction.",
    "Borrower ID, Loan Reference, Collection ID and Expense ID are the audit trail for the whole system.",
    "Blue cells are the only cells you type into. Everything grey or green calculates itself.",
    "Red flag = duplicate, invalid link or cash mismatch. Yellow flag = the row is incomplete.",
    "Dropdowns are enforced: a value that is not on this sheet is rejected on entry.",
    "Old transaction rows are history. Do not overwrite them to correct an error - add a new row.",
    "To add a collector, officer, approver or category, type it in the matching column above.",
    "Every dropdown range is sized to the list, so a new entry is picked up automatically "
    "only if it goes in the next row down - extend the range here if you add many.",
]
r = FIRST_ITEM + max(len(v[2]) for v in lists_cols.values()) + 1
ls.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
c = ls.cell(r, 1, notes[0]); c.font = F_BAND; c.fill = FILL_BAND; c.alignment = LFT
for col in range(1, 13):
    ls.cell(r, col).fill = FILL_BAND
ls.row_dimensions[r].height = 20
for i, txt in enumerate(notes[1:], start=1):
    rr = r + i
    ls.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=12)
    c = ls.cell(rr, 1, "•  " + txt)
    c.font = F_BODY; c.alignment = LFT
    ls.row_dimensions[rr].height = 16
ls.sheet_view.showGridLines = False
ls.sheet_properties.tabColor = "808080"
banner(ls, "LISTS — DROPDOWN SOURCE",
       "Every dropdown in this workbook reads from this sheet. Add names here, not in the transaction sheets.", 12)
ls.freeze_panes = "A5"

# defined names for the dropdowns
names = {key: f"Lists!${get_column_letter(ci)}${FIRST_ITEM}"
                f":${get_column_letter(ci)}${FIRST_ITEM + len(items) - 1}"
         for ci, (key, _t, items) in lists_cols.items()}
names.update({
    "BorrowerIDs":    "'Master List'!$A$9:$A$508",
    "LoanReferences": "'Loan Release'!$A$9:$A$508",
})
for n, ref in names.items():
    if n in wb.defined_names:
        del wb.defined_names[n]
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def dv_list(ws, sqref, name, title, prompt, err):
    dv = DataValidation(type="list", formula1=f"={name}", allow_blank=True,
                        showDropDown=False, showErrorMessage=True,
                        errorStyle="stop", showInputMessage=True)
    dv.promptTitle = title
    dv.prompt = prompt
    dv.errorTitle = "Value not on the list"
    dv.error = err
    dv.add(sqref)
    ws.add_data_validation(dv)
    return dv


def dv_prompt(ws, sqref, title, prompt):
    dv = DataValidation(allow_blank=True, showInputMessage=True, showErrorMessage=False)
    dv.promptTitle = title
    dv.prompt = prompt
    dv.add(sqref)
    ws.add_data_validation(dv)


def dv_number(ws, sqref, title, prompt, lo=None, hi=None, whole=False):
    dv = DataValidation(type="whole" if whole else "decimal", operator="between",
                        formula1=str(lo if lo is not None else 0),
                        formula2=str(hi if hi is not None else 999999999),
                        allow_blank=True, showErrorMessage=True, errorStyle="stop",
                        showInputMessage=True)
    dv.promptTitle = title; dv.prompt = prompt
    dv.errorTitle = "Number out of range"
    dv.error = f"Enter a number between {lo if lo is not None else 0} and {hi if hi is not None else '999,999,999'}."
    dv.add(sqref); ws.add_data_validation(dv)


def dv_date(ws, sqref, title, prompt):
    dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2000,1,1)",
                        allow_blank=True, showErrorMessage=True, errorStyle="stop",
                        showInputMessage=True)
    dv.promptTitle = title; dv.prompt = prompt
    dv.errorTitle = "Enter a real date"
    dv.error = "Type a date, e.g. 07-Apr-2026 or 4/7/2026."
    dv.add(sqref); ws.add_data_validation(dv)


# =========================================================== Master List ===
ml = wb["Master List"]
ml.delete_rows(1, ml.max_row)
banner(ml, "MASTER LIST — BORROWER REGISTER",
       "One row per borrower. Type the blue cells only; grey and green cells calculate themselves.", 14)
section(ml, "PORTFOLIO SUMMARY",
        "Live totals for the whole register. Do not type in this strip.", 14)
summary_pairs(ml, [
    ("Borrowers", "=COUNTA(A9:A508)", INT),
    ("Active", '=COUNTIF(K9:K508,"ACTIVE")', INT),
    ("Released (PHP)", "=IFERROR(SUM('Loan Release'!E9:E508),0)", MONEY),
    ("Outstanding (PHP)", "=IFERROR(SUM('Loan Release'!R9:R508),0)", MONEY),
])
ml_headers = ["Borrower ID", "Full Name", "Contact No.", "Address", "Date Registered",
              "Loan Officer", "Credit Limit (PHP)", "Loan Count", "Total Released (PHP)",
              "Total Outstanding (PHP)", "Status", "Notes", "ID Check", "Input Check"]
header_row(ml, 8, ml_headers)

put_formulas(ml, 8,  '=IF(A{r}="","",COUNTIF(\'Loan Release\'!$C$9:$C$508,A{r}))')
put_formulas(ml, 9,  '=IF(A{r}="","",IFERROR(SUMIF(\'Loan Release\'!$C$9:$C$508,A{r},\'Loan Release\'!$E$9:$E$508),0))')
put_formulas(ml, 10, '=IF(A{r}="","",IFERROR(SUMIF(\'Loan Release\'!$C$9:$C$508,A{r},\'Loan Release\'!$R$9:$R$508),0))')
put_formulas(ml, 11, '=IF(A{r}="","",IF(COUNTIFS(\'Loan Release\'!$C$9:$C$508,A{r},\'Loan Release\'!$S$9:$S$508,"PAST DUE")>0,"BLOCKED",'
                     'IF(AND(N(H{r})>0,N(J{r})=0),"INACTIVE","ACTIVE")))')
put_formulas(ml, 13, '=IF(A{r}="","",IF(COUNTIF($A$9:$A$508,A{r})>1,"DUPLICATE","OK"))')
put_formulas(ml, 14, '=IF(A{r}="","",IF(OR(B{r}="",C{r}="",D{r}="",E{r}="",F{r}=""),"INCOMPLETE","OK"))')

ml_kinds = {1: "input", 2: "input", 3: "input", 4: "input", 5: "input", 6: "input",
            7: "input", 8: "calc", 9: "calc", 10: "calc", 11: "calc", 12: "input",
            13: "flag", 14: "flag"}
style_data(ml, range(1, 15), ml_kinds,
           {1: 14, 2: 26, 3: 15, 4: 30, 5: 15, 6: 13, 7: 15, 8: 10, 9: 17, 10: 19,
            11: 12, 12: 24, 13: 13, 14: 14},
           {5: DATE, 7: MONEY, 8: INT, 9: MONEY, 10: MONEY})
for r in ROWS:
    for ci in (1, 2, 3, 4, 6, 12):
        ml.cell(r, ci).alignment = LFT
    for ci in (5, 11, 13, 14):
        ml.cell(r, ci).alignment = CTR
    for ci in (7, 9, 10):
        ml.cell(r, ci).alignment = RGT

dv_prompt(ml, "A9:A508", "Borrower ID",
          "Unique key for this borrower, e.g. BR-001. It links every loan and payment. Never reuse an ID.")
dv_list(ml, "F9:F508", "OfficerNames", "Loan Officer", "Pick from the list. Add new officers on the Lists sheet.",
        "Choose a loan officer from the dropdown, or add the name on the Lists sheet first.")
dv_number(ml, "G9:G508", "Credit Limit (PHP)", "Maximum principal this borrower may hold.", lo=0, hi=10000000)
dv_date(ml, "E9:E508", "Date Registered", "Date the borrower was registered, e.g. 07-Apr-2026.")

ml.conditional_formatting.add("A9:A508", FormulaRule(
    formula=['AND($A9<>"",COUNTIF($A$9:$A$508,$A9)>1)'],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
for col, text, fill, txt in (("M", "DUPLICATE", RED_F, RED_T),
                             ("N", "INCOMPLETE", YEL_F, YEL_T),
                             ("K", "BLOCKED", RED_F, RED_T),
                             ("K", "ACTIVE", GRN_F, GRN_T)):
    flag_cf(ml, f"{col}9:{col}508", text, fill, txt)
ml.conditional_formatting.add("J9:J508", FormulaRule(
    formula=['AND($J9<>"",$J9>0)'],
    fill=PatternFill("solid", bgColor=YEL_F), font=Font(color=YEL_T, bold=True)))
ml.freeze_panes = "C9"
ml.sheet_view.showGridLines = False
ml.sheet_properties.tabColor = "2E5A8A"
tbl = ml.tables["MasterListTable"]
tbl.ref = "A8:L508"
tbl.tableStyleInfo.name = "TableStyleMedium2"

# ------------------------------------------------------------- Loan Release =
lr = wb["Loan Release"]
lr.delete_rows(1, lr.max_row)
banner(lr, "LOAN RELEASE — ORIGINATION",
       "One row per loan. Interest is flat on the original principal; every started period counts as a full period.", 23)
section(lr, "LOAN SUMMARY", "Live totals for every loan released. Do not type in this strip.", 23)
summary_pairs(lr, [
    ("Principal Released (PHP)", "=IFERROR(SUM(E9:E508),0)", MONEY),
    ("Total Receivable (PHP)",   "=IFERROR(SUM(K9:K508),0)", MONEY),
    ("Outstanding (PHP)",        "=IFERROR(SUM(R9:R508),0)", MONEY),
    ("Active",   '=COUNTIF(S9:S508,"ACTIVE")', INT),
    ("Past Due", '=COUNTIF(S9:S508,"PAST DUE")', INT),
    ("Paid",     '=COUNTIF(S9:S508,"PAID")', INT),
])
lr_headers = ["Loan Reference", "Release Date", "Borrower ID", "Borrower Name",
              "Principal (PHP)", "Interest Rate (%)", "Interest Basis", "Interest Amount (PHP)",
              "Disbursement Channel", "Processing Fee (PHP)", "Total Receivable (PHP)",
              "Term (Days)", "Repayment Frequency", "Installment Amount (PHP)",
              "First Due Date", "Maturity Date", "Total Collected (PHP)", "Outstanding (PHP)",
              "Status", "Days Overdue", "ID Check", "Borrower Check", "Input Check"]
header_row(lr, 8, lr_headers)

put_formulas(lr, 4, '=IF(C{r}="","",IFERROR(VLOOKUP(C{r},\'Master List\'!$A$9:$B$508,2,FALSE),"NOT FOUND"))')
put_formulas(lr, 8, '=IF(OR(E{r}="",F{r}="",G{r}="",L{r}=""),"",'
                    'ROUND(E{r}*F{r}*ROUNDUP(L{r}/IF(G{r}="Daily",1,IF(G{r}="Weekly",7,'
                    'IF(G{r}="Monthly",30,IF(G{r}="Yearly",365,1)))),0),2))')
put_formulas(lr, 11, '=IF(OR(E{r}="",H{r}=""),"",ROUND(E{r}+N(H{r}),2))')
put_formulas(lr, 14, '=IF(OR(K{r}="",L{r}="",M{r}=""),"",'
                     'ROUND(K{r}/ROUNDUP(L{r}/IF(M{r}="Daily",1,IF(M{r}="Weekly",7,'
                     'IF(M{r}="Monthly",30,1))),0),2))')
put_formulas(lr, 16, '=IF(OR(O{r}="",L{r}=""),"",O{r}+L{r}-1)')
put_formulas(lr, 17, '=IF(A{r}="","",IFERROR(SUMIFS(Collection!$I$9:$I$508,Collection!$C$9:$C$508,A{r}),0))')
put_formulas(lr, 18, '=IF(NOT(ISNUMBER(K{r})),"",MAX(0,ROUND(K{r}-N(Q{r}),2)))')
put_formulas(lr, 19, '=IF(A{r}="","",IF(NOT(ISNUMBER(R{r})),"INCOMPLETE",'
                     'IF(R{r}<=0,"PAID",IF(P{r}="","ACTIVE",IF(TODAY()>P{r},"PAST DUE","ACTIVE")))))')
put_formulas(lr, 20, '=IF(A{r}="","",IF(NOT(ISNUMBER(P{r})),"",'
                     'IF(S{r}="PAID",0,MAX(0,TODAY()-P{r}))))')
put_formulas(lr, 21, '=IF(A{r}="","",IF(COUNTIF($A$9:$A$508,A{r})>1,"DUPLICATE","OK"))')
put_formulas(lr, 22, '=IF(C{r}="","",IF(D{r}="NOT FOUND","INVALID","OK"))')
put_formulas(lr, 23, '=IF(AND(A{r}="",B{r}="",C{r}=""),"",'
                     'IF(OR(A{r}="",B{r}="",C{r}="",E{r}="",F{r}="",G{r}="",I{r}="",'
                     'L{r}="",M{r}="",O{r}=""),"INCOMPLETE","OK"))')

lr_kinds = {1: "input", 2: "input", 3: "input", 4: "link", 5: "input", 6: "input",
            7: "input", 8: "calc", 9: "input", 10: "input", 11: "calc", 12: "input",
            13: "input", 14: "calc", 15: "input", 16: "calc", 17: "calc", 18: "calc",
            19: "calc", 20: "calc", 21: "flag", 22: "flag", 23: "flag"}
style_data(lr, range(1, 24), lr_kinds,
           {1: 15, 2: 13, 3: 13, 4: 24, 5: 14, 6: 12, 7: 13, 8: 15, 9: 15, 10: 14,
            11: 16, 12: 11, 13: 15, 14: 16, 15: 13, 16: 13, 17: 15, 18: 15, 19: 12,
            20: 12, 21: 12, 22: 14, 23: 13},
           {2: DATE, 5: MONEY, 6: '0.00%', 8: MONEY, 10: MONEY, 11: MONEY, 12: INT,
            14: MONEY, 15: DATE, 16: DATE, 17: MONEY, 18: MONEY, 20: INT})
for r in ROWS:
    for ci in (1, 3, 4, 9, 13, 19, 21, 22, 23):
        lr.cell(r, ci).alignment = CTR
    for ci in (5, 8, 10, 11, 14, 17, 18):
        lr.cell(r, ci).alignment = RGT
    for ci in (2, 6, 7, 12, 15, 16, 20):
        lr.cell(r, ci).alignment = CTR

dv_prompt(lr, "A9:A508", "Loan Reference",
          "Unique key for this loan, e.g. LN-001. Every payment is recorded against it. Never reuse one.")
dv_list(lr, "C9:C508", "BorrowerIDs", "Borrower ID", "Pick the borrower. The name fills in automatically.",
        "That Borrower ID is not on the Master List. Register the borrower first, then pick it here.")
dv_list(lr, "G9:G508", "InterestBasis", "Interest Basis", "The period the interest rate is quoted over.",
        "Choose Daily, Weekly, Monthly or Yearly.")
dv_list(lr, "I9:I508", "DisburseChannel", "Disbursement Channel", "How the money actually left the till.",
        "Choose Cash, GCash or Bank Transfer. Only Cash flows into the Daily Cash Report.")
dv_list(lr, "M9:M508", "RepayFrequency", "Repayment Frequency",
        "How often the borrower pays. This is separate from the interest basis.",
        "Choose Daily, Weekly or Monthly.")
dv_number(lr, "E9:E508", "Principal (PHP)", "Amount actually handed to the borrower.", lo=0, hi=10000000)
dv_number(lr, "F9:F508", "Interest Rate (%)", "Rate per interest period, e.g. 0.02 = 2%.", lo=0, hi=1)
dv_number(lr, "J9:J508", "Processing Fee (PHP)", "Fee charged on release. Not added to the receivable.", lo=0, hi=10000000)
dv_number(lr, "L9:L508", "Term (Days)", "Length of the loan in days.", lo=1, hi=3650, whole=True)
dv_date(lr, "B9:B508", "Release Date", "Date the money was released.")
dv_date(lr, "O9:O508", "First Due Date", "Date the first payment falls due.")

lr.conditional_formatting.add("A9:A508", FormulaRule(
    formula=['AND($A9<>"",COUNTIF($A$9:$A$508,$A9)>1)'],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
lr.conditional_formatting.add("E9:E508", FormulaRule(
    formula=['AND($E9<>"",$E9<=0)'],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
lr.conditional_formatting.add("F9:F508", FormulaRule(
    formula=['AND($F9<>"",OR($F9<0,$F9>1))'],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
lr.conditional_formatting.add("L9:L508", FormulaRule(
    formula=['AND($L9<>"",$L9<=0)'],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
lr.conditional_formatting.add("R9:R508", CellIsRule(
    operator="greaterThan", formula=["0"],
    fill=PatternFill("solid", bgColor=YEL_F), font=Font(color=YEL_T, bold=True)))
lr.conditional_formatting.add("T9:T508", ColorScaleRule(
    start_type="num", start_value=0, start_color="FFFFFF",
    end_type="num", end_value=60, end_color="F8696B"))
for col, text, fill, txt in (("S", "PAID", GRN_F, GRN_T),
                             ("S", "PAST DUE", RED_F, RED_T),
                             ("S", "ACTIVE", PALEBLUE, NAVY),
                             ("U", "DUPLICATE", RED_F, RED_T),
                             ("V", "INVALID", RED_F, RED_T),
                             ("W", "INCOMPLETE", YEL_F, YEL_T)):
    flag_cf(lr, f"{col}9:{col}508", text, fill, txt)
lr.freeze_panes = "D9"
lr.sheet_view.showGridLines = False
lr.sheet_properties.tabColor = "2E5A8A"
tbl = lr.tables["LoanReleaseTable"]
tbl.ref = "A8:T508"
tbl.tableColumns = tbl.tableColumns[:19] + [
    TableColumn(id=20, name="Days Overdue")]
tbl.tableStyleInfo.name = "TableStyleMedium2"

# --------------------------------------------------------------- Collection =
co = wb["Collection"]
co.delete_rows(1, co.max_row)
banner(co, "COLLECTION — PAYMENTS RECEIVED",
       "One row per payment, linked to a Loan Reference. Never overwrite a payment - add a new row.", 16)
section(co, "COLLECTION SUMMARY", "Live totals for every payment recorded. Do not type in this strip.", 16)
summary_pairs(co, [
    ("Total Collected (PHP)", "=IFERROR(SUM(I9:I508),0)", MONEY),
    ("Principal (PHP)",       "=IFERROR(SUM(F9:F508),0)", MONEY),
    ("Interest (PHP)",        "=IFERROR(SUM(G9:G508),0)", MONEY),
    ("Penalty/Other (PHP)",   "=IFERROR(SUM(H9:H508),0)", MONEY),
    ("Payments",              "=COUNT(I9:I508)", INT),
])
co_headers = ["Collection ID", "Collection Date", "Loan Reference", "Borrower ID", "Borrower Name",
              "Principal Paid (PHP)", "Interest Paid (PHP)", "Penalty/Other (PHP)",
              "Total Collected (PHP)", "Channel", "Collector", "Notes",
              "ID Check", "Loan Check", "Input Check"]
header_row(co, 8, co_headers)

put_formulas(co, 1, '=IF(OR(B{r}="",C{r}=""),"","COL-"&TEXT(B{r},"yyyymmdd")&"-"'
                    '&TEXT(COUNTIFS($B$9:B{r},B{r},$C$9:C{r},C{r}),"000"))')
put_formulas(co, 4, '=IF(C{r}="","",IFERROR(VLOOKUP(C{r},\'Loan Release\'!$A$9:$C$508,3,FALSE),"NOT FOUND"))')
put_formulas(co, 5, '=IF(C{r}="","",IFERROR(VLOOKUP(C{r},\'Loan Release\'!$A$9:$D$508,4,FALSE),"NOT FOUND"))')
put_formulas(co, 9, '=IF(COUNT(F{r}:H{r})=0,"",SUM(F{r}:H{r}))')
put_formulas(co, 13, '=IF(A{r}="","",IF(COUNTIF($A$9:$A$508,A{r})>1,"DUPLICATE","OK"))')
put_formulas(co, 14, '=IF(C{r}="","",IF(OR(D{r}="NOT FOUND",E{r}="NOT FOUND"),"INVALID","OK"))')
put_formulas(co, 15, '=IF(AND(B{r}="",C{r}=""),"",'
                     'IF(OR(B{r}="",C{r}="",F{r}="",G{r}="",J{r}=""),"INCOMPLETE","OK"))')

co_kinds = {1: "calc", 2: "input", 3: "input", 4: "link", 5: "link", 6: "input",
            7: "input", 8: "input", 9: "calc", 10: "input", 11: "input", 12: "input",
            13: "flag", 14: "flag", 15: "flag"}
style_data(co, range(1, 16), co_kinds,
           {1: 20, 2: 15, 3: 15, 4: 13, 5: 24, 6: 16, 7: 15, 8: 16, 9: 17, 10: 14,
            11: 13, 12: 24, 13: 12, 14: 12, 15: 13},
           {2: DATE, 6: MONEY, 7: MONEY, 8: MONEY, 9: MONEY})
for r in ROWS:
    for ci in (1, 2, 3, 4, 10, 11, 13, 14, 15):
        co.cell(r, ci).alignment = CTR
    for ci in (6, 7, 8, 9):
        co.cell(r, ci).alignment = RGT
    for ci in (5, 12):
        co.cell(r, ci).alignment = LFT

dv_date(co, "B9:B508", "Collection Date", "Date the money was received.")
dv_list(co, "C9:C508", "LoanReferences", "Loan Reference", "Pick the loan being repaid. Borrower ID and name fill in.",
        "That Loan Reference does not exist. Record the loan release first, then pick it here.")
dv_number(co, "F9:F508", "Principal Paid (PHP)", "Part of the payment that reduces the principal.", lo=0, hi=10000000)
dv_number(co, "G9:G508", "Interest Paid (PHP)", "Part of the payment that is interest. Enter 0 if none.", lo=0, hi=10000000)
dv_number(co, "H9:H508", "Penalty/Other (PHP)", "Late penalty or other charge. Enter 0 if none.", lo=0, hi=10000000)
dv_list(co, "J9:J508", "PaymentMethod", "Channel", "How the payment was received.",
        "Choose a channel. Only Cash flows into the Daily Cash Report.")
dv_list(co, "K9:K508", "CollectorNames", "Collector", "Who physically collected the payment.",
        "Pick a collector, or add the name on the Lists sheet first.")

co.conditional_formatting.add("C9:C508", FormulaRule(
    formula=['AND($C9<>"",COUNTIF(\'Loan Release\'!$A$9:$A$508,$C9)=0)'],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
co.conditional_formatting.add("I9:I508", FormulaRule(
    formula=['AND($I9<>"",$I9>0)'],
    fill=PatternFill("solid", bgColor=GRN_F), font=Font(color=GRN_T, bold=True)))
for col, text, fill, txt in (("M", "DUPLICATE", RED_F, RED_T),
                             ("N", "INVALID", RED_F, RED_T),
                             ("O", "INCOMPLETE", YEL_F, YEL_T),
                             ("D", "NOT FOUND", RED_F, RED_T),
                             ("E", "NOT FOUND", RED_F, RED_T)):
    flag_cf(co, f"{col}9:{col}508", text, fill, txt)
co.freeze_panes = "D9"
co.sheet_view.showGridLines = False
co.sheet_properties.tabColor = "2E5A8A"
tbl = co.tables["CollectionTable"]
tbl.ref = "A8:L508"
tbl.tableColumns = [TableColumn(id=1, name="Collection ID")] + tbl.tableColumns
tbl.tableStyleInfo.name = "TableStyleMedium2"

# ----------------------------------------------------------------- Expenses =
ex = wb["Expenses"]
ex.delete_rows(1, ex.max_row)
banner(ex, "EXPENSES — MONEY OUT",
       "One row per expense. Paid To is the recipient; Channel says whether real cash left the till.", 11)
section(ex, "EXPENSE SUMMARY", "Live totals for every expense recorded. Do not type in this strip.", 11)
summary_pairs(ex, [
    ("Total Expenses (PHP)", "=IFERROR(SUM(F9:F508),0)", MONEY),
    ("Transactions",         "=COUNT(F9:F508)", INT),
    ("Largest (PHP)",        "=IFERROR(MAX(F9:F508),0)", MONEY),
    ("Average (PHP)",        "=IFERROR(AVERAGE(F9:F508),0)", MONEY),
])
ex_headers = ["Expense ID", "Paid To", "Expense Date", "Category", "Description",
              "Amount (PHP)", "Channel", "Approved By", "Notes", "ID Check", "Input Check"]
header_row(ex, 8, ex_headers)

put_formulas(ex, 1, '=IF(OR(B{r}="",C{r}=""),"","EXP-"&TEXT(C{r},"yyyymmdd")&"-"&'
                    'TEXT(COUNTIFS($C$9:C{r},C{r},$B$9:B{r},B{r}),"000"))')
put_formulas(ex, 10, '=IF(A{r}="","",IF(COUNTIF($A$9:$A$508,A{r})>1,"DUPLICATE","OK"))')
put_formulas(ex, 11, '=IF(AND(B{r}="",C{r}=""),"",'
                     'IF(OR(B{r}="",C{r}="",D{r}="",E{r}="",F{r}="",G{r}="",H{r}=""),"INCOMPLETE","OK"))')

ex_kinds = {1: "calc", 2: "input", 3: "input", 4: "input", 5: "input", 6: "input",
            7: "input", 8: "input", 9: "input", 10: "flag", 11: "flag"}
style_data(ex, range(1, 12), ex_kinds,
           {1: 20, 2: 22, 3: 15, 4: 18, 5: 32, 6: 15, 7: 14, 8: 14, 9: 24, 10: 12, 11: 13},
           {3: DATE, 6: MONEY})
for r in ROWS:
    for ci in (1, 3, 4, 7, 8, 10, 11):
        ex.cell(r, ci).alignment = CTR
    for ci in (2, 5, 9):
        ex.cell(r, ci).alignment = LFT
    ex.cell(r, 6).alignment = RGT

dv_list(ex, "D9:D508", "ExpenseCategory", "Category", "Pick a category. Add new ones on the Lists sheet.",
        "Choose a category from the dropdown, or add it on the Lists sheet first.")
dv_list(ex, "G9:G508", "PaymentMethod", "Channel", "Cash affects the till; GCash and Bank Transfer do not.",
        "Choose a channel. Only Cash flows into the Daily Cash Report.")
dv_list(ex, "H9:H508", "ApproverNames", "Approved By", "Who authorised the spend.",
        "Pick an approver, or add the name on the Lists sheet first.")
dv_number(ex, "F9:F508", "Amount (PHP)", "Total amount of the expense.", lo=0, hi=10000000)
dv_date(ex, "C9:C508", "Expense Date", "Date the money was spent.")

for col, text, fill, txt in (("J", "DUPLICATE", RED_F, RED_T),
                             ("K", "INCOMPLETE", YEL_F, YEL_T)):
    flag_cf(ex, f"{col}9:{col}508", text, fill, txt)
ex.conditional_formatting.add("F9:F508", FormulaRule(
    formula=['AND($F9<>"",$F9>0)'],
    fill=PatternFill("solid", bgColor=YEL_F), font=Font(color=YEL_T, bold=True)))
ex.freeze_panes = "C9"
ex.sheet_view.showGridLines = False
ex.sheet_properties.tabColor = "2E5A8A"
tbl = ex.tables["ExpensesTable"]
tbl.ref = "A8:I508"
tbl.tableColumns = [TableColumn(id=1, name="Expense ID")] + tbl.tableColumns
tbl.tableStyleInfo.name = "TableStyleMedium2"

# -------------------------------------------------------- Daily Cash Report =
dc = wb["Daily Cash Report"]
dc.delete_rows(1, dc.max_row)
banner(dc, "DAILY CASH REPORT — RECONCILIATION",
       "Physical cash only. GCash and Bank Transfer never touch cash on hand, so they are excluded here.", 12)
section(dc, "CASH SUMMARY", "Live totals across every day logged. Do not type in this strip.", 12)
summary_pairs(dc, [
    ("Principal Out (PHP)",  "=IFERROR(SUM(C9:C508),0)", MONEY),
    ("Collections (PHP)",    "=IFERROR(SUM(D9:D508),0)", MONEY),
    ("Processing Fees (PHP)","=IFERROR(SUM(E9:E508),0)", MONEY),
    ("Expenses (PHP)",       "=IFERROR(SUM(F9:F508),0)", MONEY),
    ("Net Cash (PHP)",       "=IFERROR(SUM(G9:G508),0)", MONEY),
    ("Days Logged",          "=COUNT(A9:A508)", INT),
])
dc_headers = ["Date", "Beginning Cash (PHP)", "Loan Principal Out (PHP)", "Collections (PHP)",
              "Processing Fees (PHP)", "Expenses (PHP)", "Net Cash Movement (PHP)",
              "Ending Cash (PHP)", "Notes", "Actual Cash (PHP)", "Cash Variance (PHP)",
              "Reconciliation"]
header_row(dc, 8, dc_headers)

put_formulas(dc, 3, '=IF(A{r}="","",SUMIFS(\'Loan Release\'!$E$9:$E$508,\'Loan Release\'!$B$9:$B$508,A{r},'
                    '\'Loan Release\'!$I$9:$I$508,"Cash"))')
put_formulas(dc, 4, '=IF(A{r}="","",SUMIFS(Collection!$I$9:$I$508,Collection!$B$9:$B$508,A{r},'
                    'Collection!$J$9:$J$508,"Cash"))')
put_formulas(dc, 5, '=IF(A{r}="","",SUMIFS(\'Loan Release\'!$J$9:$J$508,\'Loan Release\'!$B$9:$B$508,A{r},'
                    '\'Loan Release\'!$I$9:$I$508,"Cash"))')
put_formulas(dc, 6, '=IF(A{r}="","",SUMIFS(Expenses!$F$9:$F$508,Expenses!$C$9:$C$508,A{r},'
                    'Expenses!$G$9:$G$508,"Cash"))')
put_formulas(dc, 7, '=IF(A{r}="","",SUM(D{r}:E{r})-SUM(C{r}:C{r},F{r}:F{r}))')
put_formulas(dc, 8, '=IF(A{r}="","",N(B{r})+G{r})')
put_formulas(dc, 11, '=IF(OR(A{r}="",J{r}=""),"",J{r}-H{r})')
put_formulas(dc, 12, '=IF(A{r}="","",IF(J{r}="","ENTER ACTUAL CASH",IF(ABS(K{r})<0.01,"OK","CHECK")))')
dc.cell(R0, 2, None)                      # row 9 opening balance is typed
for r in range(R0 + 1, R1 + 1):
    dc.cell(r, 2, f'=IF(A{r}="","",N(H{r-1}))')

dc_kinds = {1: "input", 2: "input", 3: "calc", 4: "calc", 5: "calc", 6: "calc",
            7: "calc", 8: "calc", 9: "input", 10: "input", 11: "calc", 12: "flag"}
style_data(dc, range(1, 13), dc_kinds,
           {1: 14, 2: 17, 3: 19, 4: 15, 5: 17, 6: 14, 7: 19, 8: 17, 9: 26, 10: 16, 11: 16, 12: 20},
           {1: DATE, 2: MONEY, 3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY,
            8: MONEY, 10: MONEY, 11: MONEY})
for r in ROWS:
    for ci in (1, 12):
        dc.cell(r, ci).alignment = CTR
    dc.cell(r, 9).alignment = LFT
    for ci in range(2, 9):
        dc.cell(r, ci).alignment = RGT
    for ci in (10, 11):
        dc.cell(r, ci).alignment = RGT
# the first day's opening balance stays a blue input cell
dc.cell(R0, 2).font = F_INPUT
dc.cell(R0, 2).fill = FILL_WHITE

dv_date(dc, "A9:A508", "Business Date", "The day this report covers. Enter days in date order.")
dv_number(dc, "B9:B9", "Opening Cash (PHP)", "Cash counted in the till at the start of the very first day logged.",
          lo=0, hi=100000000)
dv_number(dc, "J9:J508", "Actual Cash (PHP)", "Cash physically counted at the end of the day.",
          lo=0, hi=100000000)

dc.conditional_formatting.add("A10:A508", FormulaRule(
    formula=['AND($A10<>"",$A9<>"",$A10<$A9)'],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
dc.conditional_formatting.add("G9:G508", ColorScaleRule(
    start_type="min", start_color="F8696B",
    mid_type="num", mid_value=0, mid_color="FFFFFF",
    end_type="max", end_color="63BE7B"))
dc.conditional_formatting.add("H9:H508", CellIsRule(
    operator="lessThan", formula=["0"],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
dc.conditional_formatting.add("K9:K508", FormulaRule(
    formula=['AND($K9<>"",$K9<>0)'],
    fill=PatternFill("solid", bgColor=YEL_F), font=Font(color=YEL_T, bold=True)))
flag_cf(dc, "L9:L508", "CHECK", RED_F, RED_T)
flag_cf(dc, "L9:L508", "ENTER ACTUAL CASH", YEL_F, YEL_T)
flag_cf(dc, "L9:L508", "OK", GRN_F, GRN_T)
dc.freeze_panes = "B9"
dc.sheet_view.showGridLines = False
dc.sheet_properties.tabColor = "2E5A8A"
tbl = dc.tables["DailyCashTable"]
tbl.ref = "A8:I508"
tbl.tableStyleInfo.name = "TableStyleMedium2"

# ---------------------------------------------------------- Control Center =
cc = wb["Control Center"]
cc.delete_rows(1, cc.max_row)
banner(cc, "CONTROL CENTER — DATA QUALITY & RECONCILIATION",
       "Every flag in the system counted in one place. All of these should read 0 before you close the books.", 6)

checks = [
    ("Duplicate Borrower IDs",        '=COUNTIF(\'Master List\'!M9:M508,"DUPLICATE")'),
    ("Incomplete Borrower Records",   '=COUNTIF(\'Master List\'!N9:N508,"INCOMPLETE")'),
    ("Duplicate Loan References",     '=COUNTIF(\'Loan Release\'!U9:U508,"DUPLICATE")'),
    ("Invalid Loan → Borrower Links", '=COUNTIF(\'Loan Release\'!V9:V508,"INVALID")'),
    ("Incomplete Loan Records",       '=COUNTIF(\'Loan Release\'!W9:W508,"INCOMPLETE")'),
    ("Duplicate Collection IDs",      '=COUNTIF(Collection!M9:M508,"DUPLICATE")'),
    ("Invalid Collection → Loan Links", '=COUNTIF(Collection!N9:N508,"INVALID")'),
    ("Incomplete Collection Records", '=COUNTIF(Collection!O9:O508,"INCOMPLETE")'),
    ("Duplicate Expense IDs",         '=COUNTIF(Expenses!J9:J508,"DUPLICATE")'),
    ("Incomplete Expense Records",    '=COUNTIF(Expenses!K9:K508,"INCOMPLETE")'),
    ("Past Due Loans",                '=COUNTIF(\'Loan Release\'!S9:S508,"PAST DUE")'),
    ("Loans Over 30 Days Past Due",   '=COUNTIF(\'Loan Release\'!T9:T508,">30")'),
    ("Cash Reconciliation Checks",    '=COUNTIF(\'Daily Cash Report\'!L9:L508,"CHECK")'),
    ("Days Missing Actual Cash",      '=COUNTIF(\'Daily Cash Report\'!L9:L508,"ENTER ACTUAL CASH")'),
]
cc.cell(4, 1, "CONTROL").font = F_HDR
cc.cell(4, 2, "COUNT").font = F_HDR
for ci in (1, 2):
    cc.cell(4, ci).fill = FILL_HDR; cc.cell(4, ci).alignment = CTR; cc.cell(4, ci).border = BOX
FIRST = 5
for i, (label, formula) in enumerate(checks):
    r = FIRST + i
    a = cc.cell(r, 1, label); a.font = F_BODY; a.alignment = LFT; a.border = BOX
    b = cc.cell(r, 2, formula); b.font = Font(name="Calibri", size=11, bold=True)
    b.alignment = CTR; b.number_format = INT; b.border = BOX
    if i % 2:
        a.fill = PatternFill("solid", fgColor=GREY); b.fill = PatternFill("solid", fgColor=GREY)
LAST = FIRST + len(checks) - 1
cc.conditional_formatting.add(f"B{FIRST}:B{LAST}", CellIsRule(
    operator="greaterThan", formula=["0"],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
cc.conditional_formatting.add(f"B{FIRST}:B{LAST}", CellIsRule(
    operator="equal", formula=["0"],
    fill=PatternFill("solid", bgColor=GRN_F), font=Font(color=GRN_T, bold=True)))

tot = LAST + 2
a = cc.cell(tot, 1, "TOTAL OPEN ISSUES"); a.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
a.fill = FILL_BAND; a.alignment = LFT; a.border = BOX
b = cc.cell(tot, 2, f"=SUM(B{FIRST}:B{FIRST+10})+SUM(B{FIRST+12}:B{LAST})")
b.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
b.fill = FILL_BAND; b.alignment = CTR; b.number_format = INT; b.border = BOX
cc.cell(4, 1).value = "CONTROL"

r = tot + 2
cc.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
c = cc.cell(r, 1, "V9 PROCESS CONTROLS"); c.font = F_BAND; c.fill = FILL_BAND; c.alignment = LFT
for col in range(1, 7):
    cc.cell(r, col).fill = FILL_BAND
cc.row_dimensions[r].height = 20
proc = [
    ("Control Area", "What is checked", "What to do when it is flagged"),
    ("IDs", "Duplicate Borrower / Loan / Collection / Expense IDs",
     "Use a new unique ID. Never reuse an ID from an old transaction."),
    ("Links", "Every loan points at a real borrower; every payment points at a real loan",
     "Pick the correct ID from the dropdown. The dropdown only offers IDs that exist."),
    ("Required Fields", "Key fields are not blank",
     "Complete the row before treating the transaction as final."),
    ("Cash", "Actual counted cash versus expected ending cash",
     "Investigate the variance before reporting the day as reconciled."),
    ("Loan Aging", "PAST DUE status and Days Overdue",
     "Work the oldest balances first and record the collection follow-up."),
    ("History", "Old rows are retained untouched",
     "Add a new row to correct something. Never overwrite a historical transaction."),
]
for i, row_vals in enumerate(proc):
    rr = r + 1 + i
    for j, v in enumerate(row_vals):
        c = cc.cell(rr, [1, 2, 4][j], v)
        c.border = BOX
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if i == 0:
            c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTRW
        else:
            c.font = F_BODY
    cc.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    cc.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=6)
    cc.row_dimensions[rr].height = 30 if i else 22
cc.column_dimensions["A"].width = 34
cc.column_dimensions["B"].width = 12
cc.column_dimensions["C"].width = 22
cc.column_dimensions["D"].width = 24
cc.column_dimensions["E"].width = 22
cc.column_dimensions["F"].width = 22
cc.sheet_view.showGridLines = False
cc.sheet_properties.tabColor = "C00000"
cc.freeze_panes = "A5"

# ---------------------------------------------------------------- Dashboard =
db = wb["Dashboard"]
db._charts = []
db.delete_rows(1, db.max_row)
banner(db, "LENDING BUSINESS SYSTEM — EXECUTIVE DASHBOARD",
       "Everything below is calculated from Master List, Loan Release, Collection, Expenses and Daily Cash Report.", 14)

kpi = [
    (1,  "Total Borrowers",        "='Master List'!B7", INT),
    (4,  "Active Borrowers",       "='Master List'!D7", INT),
    (7,  "Principal Released",     "='Master List'!F7", MONEY),
    (10, "Outstanding",            "='Loan Release'!F7", MONEY),
]
kpi2 = [
    (1,  "Total Collected",        "=Collection!B7", MONEY),
    (4,  "Collection Rate",        "=IFERROR(Collection!B7/'Loan Release'!D7,0)", PCT),
    (7,  "Total Expenses",         "=Expenses!B7", MONEY),
    (10, "Cash on Hand (latest)",
     '=IF(COUNT(\'Daily Cash Report\'!$H$9:$H$508)=0,0,'
     'INDEX(\'Daily Cash Report\'!$H$9:$H$508,COUNT(\'Daily Cash Report\'!$H$9:$H$508)))', MONEY),
]
for band, row in ((kpi, 5), (kpi2, 8)):
    for col, label, formula, fmt in band:
        db.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 2)
        db.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 2)
        lc = db.cell(row, col, label)
        lc.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        lc.alignment = CTR
        vc = db.cell(row + 1, col, formula)
        vc.font = F_KPI_V; vc.alignment = CTR; vc.number_format = fmt
        for k in range(3):
            for rr in (row, row + 1):
                cell = db.cell(rr, col + k)
                cell.fill = FILL_BAND if rr == row else FILL_KPI
                cell.border = BOX
        db.row_dimensions[row].height = 18
        db.row_dimensions[row + 1].height = 34

def panel(ws, title, r, c1, c2):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(r, c1, title)
    cell.font = F_BAND; cell.alignment = LFT
    for k in range(c1, c2 + 1):
        ws.cell(r, k).fill = FILL_BAND
    ws.row_dimensions[r].height = 20

panel(db, "PORTFOLIO HEALTH", 12, 1, 6)
panel(db, "CASH ACTIVITY SUMMARY", 12, 8, 14)

db.cell(13, 1, "Metric"); db.cell(13, 2, "Count / Value")
db.cell(13, 4, "Metric"); db.cell(13, 5, "Count / Value")
for ci in (1, 2, 4, 5):
    c = db.cell(13, ci); c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTR; c.border = BOX

portfolio = [
    ("Active Loans",            "='Loan Release'!H7", INT),
    ("Paid Loans",              "='Loan Release'!L7", INT),
    ("Past Due Loans",          "='Loan Release'!J7", INT),
    ("Over 30 Days Past Due",   '=COUNTIF(\'Loan Release\'!T9:T508,">30")', INT),
    ("Collection Rate",         "=IFERROR(Collection!B7/'Loan Release'!D7,0)", PCT),
    ("Expense-to-Inflow",       "=IFERROR(Expenses!B7/(Collection!B7+'Daily Cash Report'!F7),0)", PCT),
    ("Outstanding Receivable",  "='Loan Release'!F7", MONEY),
]
for i, (label, formula, fmt) in enumerate(portfolio):
    r = 14 + i
    a = db.cell(r, 1, label); a.font = F_BODY; a.alignment = LFT; a.border = BOX
    b = db.cell(r, 2, formula); b.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    b.alignment = CTR; b.number_format = fmt; b.border = BOX
    if i % 2:
        a.fill = PatternFill("solid", fgColor=GREY); b.fill = PatternFill("solid", fgColor=GREY)

cash_rows = [
    ("Collections",       "='Daily Cash Report'!D7", "Loan Principal Out", "='Daily Cash Report'!B7"),
    ("Processing Fees",   "='Daily Cash Report'!F7", "Expenses",           "=Expenses!B7"),
    ("TOTAL CASH IN",     "=SUM(I14:I15)",           "TOTAL CASH OUT",     "=SUM(L14:L15)"),
]
for ci, hdr in ((8, "Cash Inflow"), (9, "Amount (PHP)"), (11, "Cash Outflow"), (12, "Amount (PHP)")):
    c = db.cell(13, ci, hdr); c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTR; c.border = BOX
for i, (lin, fin, lout, fout) in enumerate(cash_rows):
    r = 14 + i
    for col, val in ((8, lin), (9, fin), (11, lout), (12, fout)):
        c = db.cell(r, col, val)
        c.border = BOX
        if i == 2:
            c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            c.fill = FILL_BAND
        else:
            c.font = F_BODY if col in (8, 11) else Font(name="Calibri", size=11, bold=True, color=NAVY)
            if i % 2:
                c.fill = PatternFill("solid", fgColor=GREY)
        c.alignment = LFT if col in (8, 11) else RGT
        if col in (9, 12):
            c.number_format = MONEY

panel(db, "DATA QUALITY — from Control Center", 22, 1, 6)
panel(db, "DAILY CASH ACTIVITY", 22, 8, 14)
quality = [
    ("Duplicate IDs",      "='Control Center'!B5+'Control Center'!B7+'Control Center'!B10+'Control Center'!B13"),
    ("Invalid Links",      "='Control Center'!B8+'Control Center'!B11"),
    ("Incomplete Records", "='Control Center'!B6+'Control Center'!B9+'Control Center'!B12+'Control Center'!B14"),
    ("Past Due Loans",     "='Control Center'!B15"),
    ("Cash Checks",        "='Control Center'!B17"),
    ("Missing Actual Cash","='Control Center'!B18"),
]
for i, (label, formula) in enumerate(quality):
    r = 23 + i
    a = db.cell(r, 1, label); a.font = F_BODY; a.alignment = LFT; a.border = BOX
    b = db.cell(r, 2, formula); b.font = Font(name="Calibri", size=11, bold=True)
    b.alignment = CTR; b.number_format = INT; b.border = BOX
    if i % 2:
        a.fill = PatternFill("solid", fgColor=GREY); b.fill = PatternFill("solid", fgColor=GREY)
db.conditional_formatting.add("B23:B28", CellIsRule(
    operator="greaterThan", formula=["0"],
    fill=PatternFill("solid", bgColor=RED_F), font=Font(color=RED_T, bold=True)))
db.conditional_formatting.add("B23:B28", CellIsRule(
    operator="equal", formula=["0"],
    fill=PatternFill("solid", bgColor=GRN_F), font=Font(color=GRN_T, bold=True)))

# cash activity table feeding the charts
for ci, hdr in ((8, "Date"), (9, "Collections"), (10, "Principal Out"),
                (11, "Fees"), (12, "Net")):
    c = db.cell(23, ci, hdr); c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTR; c.border = BOX
for i in range(15):
    r = 24 + i
    src = R0 + i
    db.cell(r, 8, f'=IF(\'Daily Cash Report\'!A{src}="","",\'Daily Cash Report\'!A{src})').number_format = DATE
    db.cell(r, 9, f'=IF($H{r}="","",\'Daily Cash Report\'!D{src})').number_format = MONEY
    db.cell(r, 10, f'=IF($H{r}="","",\'Daily Cash Report\'!C{src})').number_format = MONEY
    db.cell(r, 11, f'=IF($H{r}="","",\'Daily Cash Report\'!E{src})').number_format = MONEY
    db.cell(r, 12, f'=IF($H{r}="","",\'Daily Cash Report\'!G{src})').number_format = MONEY
    for ci in range(8, 13):
        cc_ = db.cell(r, ci); cc_.border = BOX; cc_.font = F_LINK; cc_.fill = FILL_LINK
        cc_.alignment = CTR if ci == 8 else RGT

for col, w in {1: 26, 2: 16, 3: 3, 4: 24, 5: 16, 6: 3, 7: 3, 8: 16, 9: 16,
               10: 16, 11: 14, 12: 16, 13: 3, 14: 3}.items():
    db.column_dimensions[get_column_letter(col)].width = w
db.sheet_view.showGridLines = False
db.sheet_properties.tabColor = "1F3864"
db.freeze_panes = "A4"

line = LineChart()
line.title = "Daily Cash Movement"
line.style = 12
line.y_axis.title = "PHP"
line.x_axis.title = "Date"
line.height, line.width = 8.5, 17
data = Reference(db, min_col=9, max_col=12, min_row=23, max_row=38)
cats = Reference(db, min_col=8, min_row=24, max_row=38)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
db.add_chart(line, "A31")

bar = BarChart()
bar.type = "col"
bar.grouping = "clustered"
bar.title = "Cash In vs Cash Out"
bar.y_axis.title = "PHP"
bar.height, bar.width = 8.5, 17
bdata = Reference(db, min_col=9, max_col=11, min_row=23, max_row=38)
bar.add_data(bdata, titles_from_data=True)
bar.set_categories(cats)
db.add_chart(bar, "A49")

# ------------------------------------------------------------- Start Here ==
sh = wb.create_sheet("Start Here", 0)
banner(sh, "LENDING BUSINESS SYSTEM — START HERE",
       "ONE US MONEY · Gloria, Oriental Mindoro · Version 9", 8)
sh.sheet_view.showGridLines = False
sh.sheet_properties.tabColor = "00B050"
for col, w in {1: 4, 2: 26, 3: 22, 4: 22, 5: 22, 6: 22, 7: 22, 8: 22}.items():
    sh.column_dimensions[get_column_letter(col)].width = w

r = 4
def band(row, text, last=8):
    sh.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last)
    c = sh.cell(row, 2, text); c.font = F_BAND; c.alignment = LFT
    for k in range(2, last + 1):
        sh.cell(row, k).fill = FILL_BAND
    sh.row_dimensions[row].height = 20

def line_(row, text, col=2, bold=False, size=10, indent=0):
    sh.merge_cells(start_row=row, start_column=col + indent, end_row=row, end_column=8)
    c = sh.cell(row, col + indent, text)
    c.font = Font(name="Calibri", size=size, bold=bold, color="404040")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
    sh.row_dimensions[row].height = 17

band(r, "HOW THE SYSTEM WORKS"); r += 1
for t in ["Money moves in five steps, and each step lives on its own sheet.",
          "1.  Register the borrower on MASTER LIST. That creates a Borrower ID such as BR-001.",
          "2.  Release the loan on LOAN RELEASE. Pick the Borrower ID; the name fills in and a Loan Reference such as LN-001 identifies the loan.",
          "3.  Record every payment on COLLECTION. Pick the Loan Reference; the borrower fills in and the loan balance updates itself.",
          "4.  Record money out on EXPENSES. Only Cash-channel rows affect the till.",
          "5.  Close the day on DAILY CASH REPORT. Type the date and the counted cash; everything else is calculated.",
          "",
          "DASHBOARD is the management view. CONTROL CENTER is the error view. Neither is ever typed into."]:
    line_(r, t); r += 1
r += 1

band(r, "COLOUR CODE — what you are allowed to touch"); r += 1
legend = [("White cell, blue text", "YOU TYPE HERE. The only cells you edit.", FILL_WHITE, F_INPUT),
          ("Grey cell", "CALCULATED. Do not type - it is a formula.", FILL_CALC, F_CALC),
          ("Green cell", "LOOKED UP from another sheet. Do not type.", FILL_LINK, F_LINK),
          ("Red flag", "STOP. Duplicate ID, broken link or cash mismatch.", PatternFill("solid", fgColor=RED_F), Font(color=RED_T, bold=True)),
          ("Yellow flag", "INCOMPLETE. The row is missing a required field.", PatternFill("solid", fgColor=YEL_F), Font(color=YEL_T, bold=True)),
          ("Green flag", "OK. This row passed its checks.", PatternFill("solid", fgColor=GRN_F), Font(color=GRN_T, bold=True))]
for label, desc, fill, font in legend:
    c = sh.cell(r, 2, label); c.fill = fill; c.font = font; c.border = BOX; c.alignment = LFT
    sh.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
    d = sh.cell(r, 3, desc); d.font = F_BODY; d.alignment = LFT
    sh.row_dimensions[r].height = 18
    r += 1
r += 1

band(r, "SUGGESTED NEXT ID"); r += 1
sh.cell(r, 2, "Next Borrower ID").font = F_BODY
sh.cell(r, 2).border = BOX; sh.cell(r, 2).alignment = LFT
c = sh.cell(r, 3, '="BR-"&TEXT(COUNTA(\'Master List\'!$A$9:$A$508)+1,"000")')
c.font = Font(name="Calibri", size=12, bold=True, color=NAVY); c.fill = FILL_KPI
c.alignment = CTR; c.border = BOX
r += 1
sh.cell(r, 2, "Next Loan Reference").font = F_BODY
sh.cell(r, 2).border = BOX; sh.cell(r, 2).alignment = LFT
c = sh.cell(r, 3, '="LN-"&TEXT(COUNTA(\'Loan Release\'!$A$9:$A$508)+1,"000")')
c.font = Font(name="Calibri", size=12, bold=True, color=NAVY); c.fill = FILL_KPI
c.alignment = CTR; c.border = BOX
r += 2

band(r, "WHEN A FLAG APPEARS"); r += 1
trouble = [("Flag", "What it means", "What to do"),
           ("DUPLICATE", "Two rows share the same ID.",
            "Give the newer row a new ID. Take the next one from the box above."),
           ("INVALID", "A loan or payment points at an ID that does not exist.",
            "Re-pick the ID from the dropdown - the dropdown only offers real IDs."),
           ("INCOMPLETE", "A required field on the row is empty.",
            "Fill the blue cells. The row is not final until this clears."),
           ("ENTER ACTUAL CASH", "The day is logged but no cash was counted.",
            "Count the till and type the amount into Actual Cash on Daily Cash Report."),
           ("CHECK", "Counted cash does not match expected cash.",
            "Find the missing or mistyped transaction before closing the day."),
           ("PAST DUE", "Today is later than the maturity date and a balance remains.",
            "Work the collection. Days Overdue on Loan Release sorts worst first."),
           ("NOT FOUND", "The Borrower ID on a loan is not on the Master List.",
            "Register the borrower first, then re-pick the ID on the loan.")]
for i, row_vals in enumerate(trouble):
    for j, v in enumerate(row_vals):
        c = sh.cell(r, [2, 3, 5][j], v)
        c.border = BOX
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if i == 0:
            c.font = F_HDR; c.fill = FILL_HDR; c.alignment = CTRW
        else:
            c.font = F_BODY
    sh.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    sh.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
    sh.row_dimensions[r].height = 17 if i == 0 else 28
    r += 1
r += 1

band(r, "RULES THAT KEEP THE BOOKS AUDITABLE"); r += 1
for t in ["•  Never delete or rename an ID once it has been used.",
          "•  Never overwrite an old transaction. Add a new row instead.",
          "•  Dropdowns are enforced - a value that is not on the LISTS sheet is rejected on entry.",
          "•  To add a collector, officer, approver or expense category, type it on the LISTS sheet.",
          "•  Check CONTROL CENTER reads 0 across the board before you close the books.",
          "•  Back the file up before each month end."]:
    line_(r, t); r += 1

sh.protection.sheet = True
sh.protection.selectLockedCells = False
sh.protection.formatCells = False
sh.protection.formatColumns = False
sh.protection.formatRows = False

# ------------------------------------------------- print forms (light pass) -
for name in ("COLLECTION SHEET", "Daily Cash Report Format"):
    ws = wb[name]
    ws.sheet_properties.tabColor = "808080"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


# ------------------------------------------------ restore V8 transactions =
# The V8 file held exactly one real transaction set. It is re-entered here
# verbatim, plus the two keys that were missing and broke the relational
# chain (Loan Release!A9 and Collection!C9). Every other blank is left blank
# on purpose so the validation flags still show the real outstanding work.
from datetime import datetime as _dt

wb["Master List"]["A9"] = "BR-001"
wb["Master List"]["B9"] = "Arnel, Oroquita L"
wb["Master List"]["C9"] = "09075000487"
wb["Master List"]["D9"] = "Asiatic Ipil, Bongabong"
wb["Master List"]["E9"] = _dt(2026, 4, 7)
wb["Master List"]["F9"] = "Officer 1"

lr = wb["Loan Release"]
lr["A9"] = "LN-001"                       # <- was blank in V8: header read "Column1"
lr["B9"] = _dt(2026, 4, 7)
lr["C9"] = "BR-001"
lr["E9"] = 5000
lr["F9"] = 0.02
lr["G9"] = "Monthly"
lr["I9"] = "Cash"
lr["J9"] = 500
lr["L9"] = 30
lr["M9"] = "Daily"
lr["O9"] = _dt(2026, 4, 7)

co = wb["Collection"]
co["B9"] = _dt(2026, 4, 8)
co["C9"] = "LN-001"                       # <- was blank in V8: payment was orphaned
co["F9"] = 650

ex = wb["Expenses"]
ex["C9"] = _dt(2026, 4, 22)
ex["D9"] = "Supplies"
ex["E9"] = "Payment for Passbook"
ex["F9"] = 1500
ex["G9"] = "Cash"
ex["H9"] = "Owner"
ex["I9"] = "PAID"

wb["Daily Cash Report"]["A9"] = _dt(2026, 4, 7)
wb["Daily Cash Report"]["B9"] = 256400

# --------------------------------------------------- sheet order & options =
order = ["Start Here", "Dashboard", "Master List", "Loan Release", "Collection",
         "Expenses", "Daily Cash Report", "Control Center", "Lists",
         "COLLECTION SHEET", "Daily Cash Report Format"]
wb._sheets = [wb[n] for n in order]
wb.active = 0

for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90


# --- Excel reads column names from the TABLE definition, not the header cell.
# --- A mismatch is exactly how V8 ended up showing "Column1", so sync them.
for _ws in wb.worksheets:
    for _name, _ref in list(_ws.tables.items()):
        _t = _ws.tables[_name]
        _c1, _r1, _c2, _r2 = range_boundaries(_t.ref)
        _hdrs = [_ws.cell(_r1, _c).value for _c in range(_c1, _c2 + 1)]
        assert len(_hdrs) == len(_t.tableColumns), (_name, len(_hdrs), len(_t.tableColumns))
        assert all(_hdrs), f"{_name}: blank header in {_ws.title}"
        assert len(set(_hdrs)) == len(_hdrs), f"{_name}: duplicate header"
        for _tc, _h in zip(_t.tableColumns, _hdrs):
            _tc.name = _h
        print(f"  synced {_name:18s} <- {_ws.title} row {_r1} ({len(_hdrs)} cols)")

_n_formula = sum(1 for _w in wb.worksheets for _r in _w.iter_rows()
                 for _x in _r if isinstance(_x.value, str) and _x.value.startswith('='))

# Guard: a dropdown range with a hole in it silently offers a blank choice.
for _key, _dn in wb.defined_names.items():
    if not _dn.attr_text.startswith("Lists!"):
        continue
    _rng = _dn.attr_text.split("!")[1].replace("$", "")
    _c1, _r1, _c2, _r2 = range_boundaries(_rng)
    _vals = [ls.cell(_r, _c1).value for _r in range(_r1, _r2 + 1)]
    assert all(_vals), f"{_key} has a blank at {_dn.attr_text}: {_vals}"
    assert ls.cell(HDR_ROW, _c1).value, f"{_key} lost its header"
print(f"  dropdown guard: {sum(1 for k,d in wb.defined_names.items() if d.attr_text.startswith('Lists!'))} "
      f"lists fully populated")

# ------------------------------------------------------------ pre-save gate =
import re as _re
def _balanced(f):
    t = _re.sub(r'"[^"]*"', '""', f)
    return t.count("(") == t.count(")") and t.count('"') % 2 == 0
_broken = []
for _ws in wb.worksheets:
    for _row in _ws.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and _c.value.startswith("=") and not _balanced(_c.value):
                _broken.append(f"{_ws.title}!{_c.coordinate}")
assert not _broken, f"unbalanced formulas: {_broken[:10]} ({len(_broken)} total)"
print(f"formula gate: {_n_formula} formulas checked, all balanced")

wb.calculation.fullCalcOnLoad = True
wb.properties.title = "Lending Business System V9"
wb.properties.creator = "ONE US MONEY"
wb.properties.description = "Borrower register, loan origination, collections, expenses and daily cash reconciliation."

wb.save(DST)
print("saved", DST)
