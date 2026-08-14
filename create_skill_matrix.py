#!/usr/bin/env python3
"""Generate a ready-to-use Employee Skill Matrix Excel workbook."""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Workbook layout
# ---------------------------------------------------------------------------
NUM_SKILL_SLOTS = 16
NUM_EMPLOYEE_SLOTS = 16
NAV_ROW = 1
TITLE_ROW = 2
CATEGORY_ROW = 3
SKILL_ROW = 4
SUBHEADER_ROW = 5
DATA_START = 6
DATA_END = DATA_START + NUM_EMPLOYEE_SLOTS - 1  # 21

SHEET_HOW = "How to Use"
SHEET_MATRIX = "Skill Matrix"
SHEET_DASH = "Dashboard"
SHEET_EMP = "Employees"
SHEET_SKILLS = "Skills"
SHEET_SETTINGS = "Settings"
SHEET_DATA = "Data"

NAV_SHEETS = [
    SHEET_HOW,
    SHEET_MATRIX,
    SHEET_DASH,
    SHEET_EMP,
    SHEET_SKILLS,
    SHEET_SETTINGS,
]

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
NAVY = "1B365D"
TEAL = "0E7C7B"
TEAL_DARK = "0A5C5B"
GOLD = "C4A35A"
WHITE = "FFFFFF"
INK = "1F2933"
MUTED = "5E6C7A"
LINE = "D0D7DE"
PAPER = "F4F6F8"
LY_FILL = "E8F1FA"
TY_FILL = "E6F6F1"
CHANGE_FILL = "F7F5F2"
GREEN = "1E7A46"
RED = "B42318"
AMBER = "B45309"

RATING_FILLS = {
    1: "F5D0D0",
    2: "FAD9C4",
    3: "FBF3D0",
    4: "D4EDDA",
    5: "A8D5BA",
}

CATEGORY_COLORS = {
    "Technical": "1B365D",
    "Tools": "3D5A80",
    "Soft Skills": "0E7C7B",
    "Domain": "6B5B95",
}

THIN = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
MED = Border(
    left=Side(style="medium", color=NAVY),
    right=Side(style="medium", color=NAVY),
    top=Side(style="medium", color=NAVY),
    bottom=Side(style="medium", color=NAVY),
)

FONT = "Calibri"

# ---------------------------------------------------------------------------
# Sample data (replace with your team)
# ---------------------------------------------------------------------------
THIS_YEAR = 2026
LAST_YEAR = 2025
TARGET = 3

EMPLOYEES = [
    ("EMP-001", "Alex Rivera", "Engineering", "Software Engineer", "Dana Wright"),
    ("EMP-002", "Jordan Chen", "Engineering", "Senior Developer", "Dana Wright"),
    ("EMP-003", "Sam Patel", "Data", "Data Analyst", "Quinn Foster"),
    ("EMP-004", "Taylor Brooks", "Product", "Product Manager", "Quinn Foster"),
    ("EMP-005", "Morgan Lee", "Design", "UX Designer", "Quinn Foster"),
    ("EMP-006", "Casey Nguyen", "Engineering", "QA Engineer", "Dana Wright"),
    ("EMP-007", "Riley Thompson", "Operations", "Operations Lead", "Dana Wright"),
    ("EMP-008", "Avery Kim", "Data", "Data Engineer", "Quinn Foster"),
]

SKILLS = [
    ("SK-001", "Python", "Technical", "Writing, reading, and debugging Python code"),
    ("SK-002", "SQL", "Technical", "Querying and transforming data in databases"),
    ("SK-003", "Excel", "Technical", "Spreadsheets, formulas, pivot tables, and reporting"),
    ("SK-004", "Data Analysis", "Technical", "Interpreting data to support decisions"),
    ("SK-005", "Git", "Tools", "Version control, branching, and code review"),
    ("SK-006", "Communication", "Soft Skills", "Clear writing, speaking, and listening"),
    ("SK-007", "Leadership", "Soft Skills", "Guiding others and owning outcomes"),
    ("SK-008", "Problem Solving", "Soft Skills", "Breaking down issues and finding solutions"),
    ("SK-009", "Project Management", "Domain", "Planning, tracking, and delivering work"),
    ("SK-010", "Stakeholder Management", "Domain", "Aligning partners and managing expectations"),
]

# (last year, this year) for each employee row × skill column
RATINGS = [
    [(3, 4), (3, 3), (2, 3), (2, 3), (4, 4), (3, 4), (2, 2), (4, 4), (2, 2), (2, 3)],
    [(5, 5), (4, 4), (3, 3), (3, 4), (5, 5), (4, 4), (3, 4), (5, 5), (3, 3), (3, 3)],
    [(2, 3), (5, 5), (5, 5), (4, 5), (2, 3), (4, 4), (2, 3), (4, 4), (3, 3), (3, 4)],
    [(1, 2), (2, 2), (4, 4), (3, 3), (1, 2), (5, 5), (4, 5), (4, 4), (5, 5), (5, 5)],
    [(1, 1), (1, 2), (3, 4), (2, 3), (2, 2), (5, 5), (3, 3), (4, 5), (3, 3), (4, 4)],
    [(2, 3), (3, 3), (4, 4), (3, 3), (4, 4), (3, 4), (2, 2), (4, 5), (2, 3), (2, 3)],
    [(1, 1), (2, 2), (5, 5), (3, 3), (1, 1), (4, 4), (4, 4), (4, 4), (4, 5), (4, 4)],
    [(4, 5), (5, 5), (3, 3), (5, 5), (4, 5), (3, 3), (2, 3), (4, 4), (2, 3), (2, 3)],
]

RATING_SCALE = [
    (1, "Beginner", "Little or no hands-on experience. Needs close guidance."),
    (2, "Basic", "Can complete simple tasks with help or a checklist."),
    (3, "Competent", "Works independently on standard tasks. Meets the target."),
    (4, "Proficient", "Handles complex work and can coach others."),
    (5, "Expert", "Go-to specialist. Sets the standard for the team."),
]


def fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def font(size=11, bold=False, color=INK, italic=False, underline=None) -> Font:
    return Font(
        name=FONT,
        size=size,
        bold=bold,
        italic=italic,
        color=color,
        underline=underline,
    )


def align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def skill_columns(n: int) -> tuple[int, int, int]:
    """1-based (last year, this year, change) columns for skill slot n (1-based)."""
    start = 3 + (n - 1) * 3
    return start, start + 1, start + 2


def matrix_last_col() -> int:
    return 2 + NUM_SKILL_SLOTS * 3


def q(sheet: str) -> str:
    return f"'{sheet}'"


def add_defined_name(wb: Workbook, name: str, ref: str) -> None:
    wb.defined_names.add(DefinedName(name=name, attr_text=ref))


def add_navigation(ws, active: str, last_col: int) -> None:
    for i, name in enumerate(NAV_SHEETS, start=1):
        cell = ws.cell(NAV_ROW, i, name)
        cell.hyperlink = f"#'{name}'!A1"
        cell.font = font(
            11,
            bold=name == active,
            color=WHITE,
            underline="single",
        )
        cell.fill = fill(TEAL if name == active else NAVY)
        cell.alignment = align("center")
        cell.border = Border(
            left=Side(style="thin", color="0A3A5C"),
            right=Side(style="thin", color="0A3A5C"),
            top=Side(style="thin", color="0A3A5C"),
            bottom=Side(style="thin", color="0A3A5C"),
        )
    for col in range(len(NAV_SHEETS) + 1, last_col + 1):
        cell = ws.cell(NAV_ROW, col, "")
        cell.fill = fill(NAVY)
    ws.row_dimensions[NAV_ROW].height = 24


def style_title(cell, text: str, size: int = 22) -> None:
    cell.value = text
    cell.font = font(size, bold=True, color=NAVY)
    cell.alignment = align("left", "center")


def apply_print(ws, landscape=True, fit_width=1, fit_height=1) -> None:
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = fit_width
    ws.page_setup.fitToHeight = fit_height
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "&BEmployee Skill Matrix"
    ws.oddFooter.right.text = "Page &P of &N"


def add_table(ws, name: str, ref: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


# ===========================================================================
# Settings
# ===========================================================================
def build_settings(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_SETTINGS)
    last_col = 8
    add_navigation(ws, SHEET_SETTINGS, last_col)
    ws.merge_cells("A2:H2")
    style_title(ws["A2"], "Settings")
    ws["A3"] = "Change the years and target rating here. Headers and the dashboard update automatically."
    ws["A3"].font = font(11, color=MUTED, italic=True)
    ws.merge_cells("A3:H3")

    ws["A5"] = "Label"
    ws["B5"] = "Value"
    ws["C5"] = "What it controls"
    for col in range(1, 4):
        ws.cell(5, col).font = font(11, bold=True, color=WHITE)
        ws.cell(5, col).fill = fill(NAVY)
        ws.cell(5, col).alignment = align("center")

    rows = [
        (6, "This year", THIS_YEAR, "Appears as the This Year column on the Skill Matrix and Dashboard."),
        (7, "Last year", LAST_YEAR, "Appears as the Last Year column for comparison."),
        (8, "Minimum rating", 1, "Lowest score you can enter. Keep at 1."),
        (9, "Maximum rating", 5, "Highest score you can enter. Keep at 5."),
        (10, "Target rating", TARGET, "Scores below this are flagged as development needs."),
    ]
    for r, label, value, note in rows:
        ws.cell(r, 1, label).font = font(11, bold=True)
        ws.cell(r, 1).fill = fill(PAPER)
        ws.cell(r, 1).border = THIN
        cell = ws.cell(r, 2, value)
        cell.font = font(14, bold=True, color=TEAL)
        cell.alignment = align("center")
        cell.fill = fill("FFF8E7")
        cell.border = MED
        ws.cell(r, 3, note).font = font(11, color=MUTED)
        ws.cell(r, 3).border = THIN
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)

    add_defined_name(wb, "ThisYear", f"{q(SHEET_SETTINGS)}!$B$6")
    add_defined_name(wb, "LastYear", f"{q(SHEET_SETTINGS)}!$B$7")
    add_defined_name(wb, "MinRating", f"{q(SHEET_SETTINGS)}!$B$8")
    add_defined_name(wb, "MaxRating", f"{q(SHEET_SETTINGS)}!$B$9")
    add_defined_name(wb, "TargetRating", f"{q(SHEET_SETTINGS)}!$B$10")

    ws["A12"] = "Rating scale (1–5)"
    ws["A12"].font = font(14, bold=True, color=NAVY)
    ws.merge_cells("A12:C12")
    headers = ["Rating", "Level", "Meaning"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(13, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")
        c.border = THIN
    for i, (score, level, meaning) in enumerate(RATING_SCALE):
        r = 14 + i
        c0 = ws.cell(r, 1, score)
        c0.font = font(16, bold=True, color=INK)
        c0.fill = fill(RATING_FILLS[score])
        c0.alignment = align("center")
        c0.border = THIN
        c1 = ws.cell(r, 2, level)
        c1.font = font(12, bold=True)
        c1.fill = fill(RATING_FILLS[score])
        c1.border = THIN
        c2 = ws.cell(r, 3, meaning)
        c2.font = font(11)
        c2.fill = fill(RATING_FILLS[score])
        c2.border = THIN
        c2.alignment = align("left", wrap=True)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 28

    ws["A20"] = "Starting a new review year"
    ws["A20"].font = font(14, bold=True, color=NAVY)
    ws.merge_cells("A20:H20")
    steps = [
        "1. Copy every This Year rating on the Skill Matrix and paste it into the matching Last Year column (Paste Values).",
        "2. Clear the This Year columns, then enter the new ratings.",
        "3. Update This year and Last year in the yellow cells above.",
        "4. The Change columns and Dashboard will recalculate on their own.",
    ]
    for i, text in enumerate(steps):
        cell = ws.cell(21 + i, 1, text)
        cell.font = font(11)
        cell.alignment = align("left", wrap=True)
        ws.merge_cells(start_row=21 + i, start_column=1, end_row=21 + i, end_column=8)
        ws.row_dimensions[21 + i].height = 22

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    for col in range(3, 9):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A2"
    ws.sheet_properties.tabColor = GOLD
    apply_print(ws, landscape=False, fit_height=1)


# ===========================================================================
# Employees
# ===========================================================================
def build_employees(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_EMP)
    last_col = 6
    add_navigation(ws, SHEET_EMP, last_col)
    ws.merge_cells("A2:F2")
    style_title(ws["A2"], "Employees")
    ws.merge_cells("A3:F3")
    ws["A3"] = (
        "Add a person in the next empty row (keep the Employee ID). "
        "They will appear automatically on the Skill Matrix. "
        f"{NUM_EMPLOYEE_SLOTS} slots are set up; insert a row in the table if you need more."
    )
    ws["A3"].font = font(11, color=MUTED, italic=True)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 36

    headers = ["Employee ID", "Full Name", "Department", "Job Title", "Manager", "Status"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")

    for i in range(NUM_EMPLOYEE_SLOTS):
        r = 6 + i
        emp_id = f"EMP-{i + 1:03d}"
        if i < len(EMPLOYEES):
            _, name, dept, title, manager = EMPLOYEES[i]
            status = "Active"
        else:
            name = dept = title = manager = ""
            status = ""
        values = [emp_id, name, dept, title, manager, status]
        for c, val in enumerate(values, 1):
            cell = ws.cell(r, c, val)
            cell.font = font(11)
            cell.border = THIN
            cell.alignment = align("left" if c > 1 else "center")
            if i >= len(EMPLOYEES):
                cell.fill = fill("FFF8E7")

    first_empty = 6 + len(EMPLOYEES)
    ws.cell(first_empty, 2).comment = Comment(
        "Type a new employee's full name here. They will appear on the Skill Matrix automatically.",
        "Skill Matrix",
        width=240,
        height=80,
    )

    add_table(ws, "tblEmployees", f"A5:F{5 + NUM_EMPLOYEE_SLOTS}")
    widths = [14, 22, 16, 22, 18, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = "3D5A80"
    apply_print(ws, landscape=False)


# ===========================================================================
# Skills
# ===========================================================================
def build_skills(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_SKILLS)
    last_col = 4
    add_navigation(ws, SHEET_SKILLS, last_col)
    ws.merge_cells("A2:D2")
    style_title(ws["A2"], "Skills catalog")
    ws.merge_cells("A3:D3")
    ws["A3"] = (
        "To add a skill later: type a name in the next yellow row. "
        "The Skill Matrix picks it up in the next empty column group. "
        "Then enter Last Year and This Year ratings (1–5) for each employee."
    )
    ws["A3"].font = font(11, color=MUTED, italic=True)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 48

    headers = ["Skill ID", "Skill Name", "Category", "Description"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")

    for i in range(NUM_SKILL_SLOTS):
        r = 6 + i
        skill_id = f"SK-{i + 1:03d}"
        if i < len(SKILLS):
            _, name, category, desc = SKILLS[i]
        else:
            name = category = desc = ""
        values = [skill_id, name, category, desc]
        for c, val in enumerate(values, 1):
            cell = ws.cell(r, c, val)
            cell.font = font(11)
            cell.border = THIN
            cell.alignment = align("left" if c > 1 else "center", wrap=True)
            if i >= len(SKILLS):
                cell.fill = fill("FFF8E7")
        ws.row_dimensions[r].height = 22

    # Category dropdown for the catalog
    dv = DataValidation(
        type="list",
        formula1='"Technical,Tools,Soft Skills,Domain"',
        allow_blank=True,
    )
    dv.error = "Choose Technical, Tools, Soft Skills, or Domain — or add a new category to this list."
    dv.errorTitle = "Category"
    dv.prompt = "Optional: group related skills"
    dv.promptTitle = "Category"
    ws.add_data_validation(dv)
    dv.add(f"C6:C{5 + NUM_SKILL_SLOTS}")

    first_empty = 6 + len(SKILLS)
    ws.cell(first_empty, 2).comment = Comment(
        "Type a new skill name here. It becomes the next column group on the Skill Matrix.",
        "Skill Matrix",
        width=240,
        height=80,
    )

    add_table(ws, "tblSkills", f"A5:D{5 + NUM_SKILL_SLOTS}")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 62
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = TEAL
    apply_print(ws, landscape=False)


# ===========================================================================
# Skill Matrix
# ===========================================================================
def build_matrix(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_MATRIX, 0)
    last_col = matrix_last_col()
    add_navigation(ws, SHEET_MATRIX, last_col)

    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=min(8, last_col))
    style_title(ws["A2"], "Employee skill matrix")
    ws.merge_cells(start_row=TITLE_ROW, start_column=9, end_row=TITLE_ROW, end_column=last_col)
    hint = ws.cell(
        TITLE_ROW,
        9,
        'Rate 1 (beginner) to 5 (expert). Yellow skill rows on the Skills sheet become new columns here. Do not type in Change — it is calculated.',
    )
    hint.font = font(11, italic=True, color=MUTED)
    hint.alignment = align("left", wrap=True)
    ws.row_dimensions[TITLE_ROW].height = 32

    # Employee / department headers (rows 3-5)
    ws.merge_cells(start_row=CATEGORY_ROW, start_column=1, end_row=SUBHEADER_ROW, end_column=1)
    ws.merge_cells(start_row=CATEGORY_ROW, start_column=2, end_row=SUBHEADER_ROW, end_column=2)
    for col, label in ((1, "Employee"), (2, "Department")):
        cell = ws.cell(CATEGORY_ROW, col, label)
        cell.font = font(12, bold=True, color=WHITE)
        cell.fill = fill(NAVY)
        cell.alignment = align("center", wrap=True)
        cell.border = THIN
        ws.cell(SKILL_ROW, col).fill = fill(NAVY)
        ws.cell(SKILL_ROW, col).border = THIN
        ws.cell(SUBHEADER_ROW, col).fill = fill(NAVY)
        ws.cell(SUBHEADER_ROW, col).border = THIN

    # Skill column groups
    for n in range(1, NUM_SKILL_SLOTS + 1):
        ly, ty, ch = skill_columns(n)
        skills_row = 5 + n  # Skills sheet data starts at row 6, slot 1 -> row 6
        # Category (formula)
        ws.merge_cells(start_row=CATEGORY_ROW, start_column=ly, end_row=CATEGORY_ROW, end_column=ch)
        cat = ws.cell(
            CATEGORY_ROW,
            ly,
            f'=IF({q(SHEET_SKILLS)}!B{skills_row}="","",{q(SHEET_SKILLS)}!C{skills_row})',
        )
        cat.font = font(9, bold=True, color=WHITE)
        cat.fill = fill(TEAL)
        cat.alignment = align("center")
        cat.border = THIN
        for col in range(ly, ch + 1):
            ws.cell(CATEGORY_ROW, col).fill = fill(TEAL)
            ws.cell(CATEGORY_ROW, col).border = THIN

        # Skill name
        ws.merge_cells(start_row=SKILL_ROW, start_column=ly, end_row=SKILL_ROW, end_column=ch)
        name = ws.cell(
            SKILL_ROW,
            ly,
            f'=IF({q(SHEET_SKILLS)}!B{skills_row}="","",{q(SHEET_SKILLS)}!B{skills_row})',
        )
        name.font = font(11, bold=True, color=NAVY)
        name.fill = fill(PAPER)
        name.alignment = align("center", wrap=True)
        name.border = THIN
        for col in range(ly, ch + 1):
            ws.cell(SKILL_ROW, col).fill = fill(PAPER)
            ws.cell(SKILL_ROW, col).border = THIN

        # Subheaders with year from Settings
        ly_h = ws.cell(SUBHEADER_ROW, ly, f'=LastYear')
        ty_h = ws.cell(SUBHEADER_ROW, ty, f'=ThisYear')
        ch_h = ws.cell(SUBHEADER_ROW, ch, "Change")
        ly_h.fill = fill("5B8FB9")
        ty_h.fill = fill(TEAL)
        ch_h.fill = fill("8A8178")
        for cell in (ly_h, ty_h, ch_h):
            cell.font = font(10, bold=True, color=WHITE)
            cell.alignment = align("center")
            cell.border = THIN

        ws.column_dimensions[get_column_letter(ly)].width = 11
        ws.column_dimensions[get_column_letter(ty)].width = 11
        ws.column_dimensions[get_column_letter(ch)].width = 10

    # Data rows
    for i in range(NUM_EMPLOYEE_SLOTS):
        r = DATA_START + i
        emp_row = 6 + i
        name_cell = ws.cell(
            r,
            1,
            f'=IF({q(SHEET_EMP)}!B{emp_row}="","",{q(SHEET_EMP)}!B{emp_row})',
        )
        dept_cell = ws.cell(
            r,
            2,
            f'=IF({q(SHEET_EMP)}!B{emp_row}="","",{q(SHEET_EMP)}!C{emp_row})',
        )
        name_cell.font = font(11, bold=True)
        dept_cell.font = font(11, color=MUTED)
        name_cell.alignment = align("left")
        dept_cell.alignment = align("left")
        name_cell.border = THIN
        dept_cell.border = THIN
        name_cell.fill = fill(WHITE)
        dept_cell.fill = fill(WHITE)
        ws.row_dimensions[r].height = 22

        for n in range(1, NUM_SKILL_SLOTS + 1):
            ly, ty, ch = skill_columns(n)
            ly_cell = ws.cell(r, ly, "")
            ty_cell = ws.cell(r, ty, "")
            ch_cell = ws.cell(
                r,
                ch,
                f'=IF(OR({get_column_letter(ly)}{r}="",{get_column_letter(ty)}{r}=""),"",'
                f"{get_column_letter(ty)}{r}-{get_column_letter(ly)}{r})",
            )
            ly_cell.fill = fill(LY_FILL)
            ty_cell.fill = fill(TY_FILL)
            ch_cell.fill = fill(CHANGE_FILL)
            ly_cell.alignment = align("center")
            ty_cell.alignment = align("center")
            ch_cell.alignment = align("center")
            ly_cell.border = THIN
            ty_cell.border = THIN
            ch_cell.border = THIN
            ly_cell.font = font(12, bold=True)
            ty_cell.font = font(12, bold=True)
            ch_cell.font = font(11)
            ch_cell.number_format = "+0;-0;0"

            emp_i = i
            skill_i = n - 1
            if emp_i < len(RATINGS) and skill_i < len(RATINGS[0]):
                last_year, this_year = RATINGS[emp_i][skill_i]
                ly_cell.value = last_year
                ty_cell.value = this_year

    # Data validation 1-5
    dv = DataValidation(
        type="whole",
        operator="between",
        formula1="1",
        formula2="5",
        allow_blank=True,
        showErrorMessage=True,
        showInputMessage=True,
    )
    dv.errorTitle = "Rating must be 1–5"
    dv.error = "Enter a whole number from 1 (beginner) to 5 (expert), or leave the cell blank."
    dv.promptTitle = "Skill rating"
    dv.prompt = "1 = beginner · 2 = basic · 3 = competent · 4 = proficient · 5 = expert"
    ws.add_data_validation(dv)

    for n in range(1, NUM_SKILL_SLOTS + 1):
        ly, ty, _ = skill_columns(n)
        dv.add(f"{get_column_letter(ly)}{DATA_START}:{get_column_letter(ty)}{DATA_END}")

        ly_range = f"{get_column_letter(ly)}{DATA_START}:{get_column_letter(ly)}{DATA_END}"
        ty_range = f"{get_column_letter(ty)}{DATA_START}:{get_column_letter(ty)}{DATA_END}"
        ch_range = f"{get_column_letter(skill_columns(n)[2])}{DATA_START}:{get_column_letter(skill_columns(n)[2])}{DATA_END}"

        for rng in (ly_range, ty_range):
            for score, color in RATING_FILLS.items():
                ws.conditional_formatting.add(
                    rng,
                    CellIsRule(
                        operator="equal",
                        formula=[str(score)],
                        fill=fill(color),
                        font=font(12, bold=True, color=INK),
                    ),
                )

        ws.conditional_formatting.add(
            ch_range,
            CellIsRule(
                operator="greaterThan",
                formula=["0"],
                fill=fill("D4EDDA"),
                font=font(11, bold=True, color=GREEN),
            ),
        )
        ws.conditional_formatting.add(
            ch_range,
            CellIsRule(
                operator="lessThan",
                formula=["0"],
                fill=fill("F8D7DA"),
                font=font(11, bold=True, color=RED),
            ),
        )
        ws.conditional_formatting.add(
            ch_range,
            CellIsRule(
                operator="equal",
                formula=["0"],
                fill=fill("E9ECEF"),
                font=font(11, color=MUTED),
            ),
        )

    # Category header color by value
    for n in range(1, NUM_SKILL_SLOTS + 1):
        ly, _, ch = skill_columns(n)
        rng = f"{get_column_letter(ly)}{CATEGORY_ROW}:{get_column_letter(ch)}{CATEGORY_ROW}"
        for category, color in CATEGORY_COLORS.items():
            ws.conditional_formatting.add(
                rng,
                FormulaRule(
                    formula=[f'{get_column_letter(ly)}{CATEGORY_ROW}="{category}"'],
                    fill=fill(color),
                    font=font(9, bold=True, color=WHITE),
                ),
            )
        ws.conditional_formatting.add(
            rng,
            FormulaRule(
                formula=[f'{get_column_letter(ly)}{CATEGORY_ROW}=""'],
                fill=fill("9AA5B1"),
                font=font(9, bold=True, color=WHITE),
            ),
        )
        # Unused skill placeholder styling
        name_rng = f"{get_column_letter(ly)}{SKILL_ROW}:{get_column_letter(ch)}{SKILL_ROW}"
        ws.conditional_formatting.add(
            name_rng,
            FormulaRule(
                formula=[f'{q(SHEET_SKILLS)}!B{5 + n}=""'],
                fill=fill("FFF8E7"),
                font=font(10, italic=True, color=MUTED),
            ),
        )

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    ws.row_dimensions[CATEGORY_ROW].height = 18
    ws.row_dimensions[SKILL_ROW].height = 32
    ws.row_dimensions[SUBHEADER_ROW].height = 20
    # Legend under the grid so raters do not have to leave this sheet
    legend_row = DATA_END + 2
    ws.merge_cells(start_row=legend_row, start_column=1, end_row=legend_row, end_column=2)
    ws.cell(legend_row, 1, "Rating scale").font = font(12, bold=True, color=NAVY)
    for i, (score, level, meaning) in enumerate(RATING_SCALE):
        col = 3 + i
        cell = ws.cell(legend_row, col, f"{score}  {level}")
        cell.fill = fill(RATING_FILLS[score])
        cell.font = font(10, bold=True)
        cell.alignment = align("center")
        cell.border = THIN
    ws.merge_cells(start_row=legend_row + 1, start_column=1, end_row=legend_row + 1, end_column=8)
    foot = ws.cell(
        legend_row + 1,
        1,
        "To add a skill: type its name in the next yellow row on the Skills sheet. "
        "To add a person: use the next yellow row on the Employees sheet. "
        "Green Change = improved vs last year; red = declined.",
    )
    foot.font = font(10, italic=True, color=MUTED)
    foot.alignment = align("left", wrap=True)
    ws.row_dimensions[legend_row + 1].height = 32

    ws.freeze_panes = f"C{DATA_START}"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.zoomScale = 90
    apply_print(ws, landscape=True, fit_width=1, fit_height=1)
    ws.page_setup.fitToHeight = 1
    ws.print_title_rows = "1:5"
    ws.print_title_cols = "A:B"


# ===========================================================================
# Hidden Data sheet (unpivoted ratings for dashboard formulas)
# ===========================================================================
def build_data(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_DATA)
    headers = [
        "Employee",
        "Department",
        "Skill",
        "Category",
        "Last Year",
        "This Year",
        "Change",
    ]
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)

    row = 2
    for emp_i in range(NUM_EMPLOYEE_SLOTS):
        matrix_row = DATA_START + emp_i
        for skill_i in range(NUM_SKILL_SLOTS):
            ly, ty, ch = skill_columns(skill_i + 1)
            emp_ref = f"{q(SHEET_MATRIX)}!A{matrix_row}"
            skill_ref = f"{q(SHEET_MATRIX)}!{get_column_letter(ly)}{SKILL_ROW}"
            occupied = f'AND({emp_ref}<>"",{q(SHEET_SKILLS)}!B{6 + skill_i}<>"")'
            ws.cell(row, 1, f'=IF({occupied},{emp_ref},"")')
            ws.cell(row, 2, f'=IF(A{row}="","",{q(SHEET_MATRIX)}!B{matrix_row})')
            ws.cell(row, 3, f'=IF(A{row}="","",{skill_ref})')
            ws.cell(
                row,
                4,
                f'=IF(A{row}="","",{q(SHEET_MATRIX)}!{get_column_letter(ly)}{CATEGORY_ROW})',
            )
            ws.cell(
                row,
                5,
                f'=IF(OR(A{row}="",{q(SHEET_MATRIX)}!{get_column_letter(ly)}{matrix_row}=""),"",'
                f"{q(SHEET_MATRIX)}!{get_column_letter(ly)}{matrix_row})",
            )
            ws.cell(
                row,
                6,
                f'=IF(OR(A{row}="",{q(SHEET_MATRIX)}!{get_column_letter(ty)}{matrix_row}=""),"",'
                f"{q(SHEET_MATRIX)}!{get_column_letter(ty)}{matrix_row})",
            )
            ws.cell(row, 7, f'=IF(OR(E{row}="",F{row}=""),"",F{row}-E{row})')
            row += 1

    last_data_row = row - 1
    add_defined_name(wb, "DataLastYear", f"{q(SHEET_DATA)}!$E$2:$E${last_data_row}")
    add_defined_name(wb, "DataThisYear", f"{q(SHEET_DATA)}!$F$2:$F${last_data_row}")
    add_defined_name(wb, "DataSkill", f"{q(SHEET_DATA)}!$C$2:$C${last_data_row}")
    add_defined_name(wb, "DataEmployee", f"{q(SHEET_DATA)}!$A$2:$A${last_data_row}")
    add_defined_name(wb, "DataChange", f"{q(SHEET_DATA)}!$G$2:$G${last_data_row}")

    ws.sheet_state = "hidden"
    for col, w in enumerate([22, 16, 24, 16, 12, 12, 12], 1):
        ws.column_dimensions[get_column_letter(col)].width = w


# ===========================================================================
# Dashboard
# ===========================================================================
def build_dashboard(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_DASH, 1)
    last_col = 16
    add_navigation(ws, SHEET_DASH, last_col)

    ws.merge_cells("A2:H2")
    style_title(ws["A2"], "Year-over-year dashboard")
    ws.merge_cells("A3:H3")
    ws["A3"] = (
        f'=CONCATENATE("Comparing ",LastYear," with ",ThisYear,'
        f'". Target rating: ",TargetRating,". Add skills or people and this page updates automatically.")'
    )
    ws["A3"].font = font(11, italic=True, color=MUTED)

    # KPI cards
    cards = [
        (1, "This year average", "=IFERROR(ROUND(AVERAGE(DataThisYear),2),\"—\")", TEAL, "0.00"),
        (4, "Last year average", "=IFERROR(ROUND(AVERAGE(DataLastYear),2),\"—\")", "5B8FB9", "0.00"),
        (7, "Year-over-year", "=IFERROR(ROUND(AVERAGE(DataThisYear)-AVERAGE(DataLastYear),2),\"—\")", GOLD, "+0.00;-0.00;0.00"),
        (10, "Skills below target", "=COUNTIF(H12:H27,\"Yes\")", RED, "0"),
        (13, "Ratings entered", "=COUNT(DataThisYear)", NAVY, "0"),
    ]
    for col, title, formula, color, fmt in cards:
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 2)
        ws.merge_cells(start_row=6, start_column=col, end_row=7, end_column=col + 2)
        head = ws.cell(5, col, title)
        head.font = font(10, bold=True, color=WHITE)
        head.fill = fill(color)
        head.alignment = align("center")
        val = ws.cell(6, col, formula)
        val.font = font(26, bold=True, color=NAVY if color == GOLD else WHITE)
        val.fill = fill("FFF8E7" if color == GOLD else color)
        val.alignment = align("center")
        val.number_format = fmt
        for r in (5, 6, 7):
            for c in range(col, col + 3):
                ws.cell(r, c).fill = fill("FFF8E7" if color == GOLD and r > 5 else color)
                ws.cell(r, c).border = Border(
                    left=Side(style="thin", color=WHITE),
                    right=Side(style="thin", color=WHITE),
                    top=Side(style="thin", color=WHITE),
                    bottom=Side(style="thin", color=WHITE),
                )
        ws.cell(7, col).fill = fill("FFF8E7" if color == GOLD else color)

    # Skill comparison table
    ws.merge_cells("A10:H10")
    ws["A10"] = "Skill comparison — last year vs this year"
    ws["A10"].font = font(14, bold=True, color=NAVY)

    skill_headers = [
        "Skill",
        "Category",
        "Last year avg",
        "This year avg",
        "Change",
        "Status",
        "# rated",
        "Below target?",
    ]
    for i, h in enumerate(skill_headers, 1):
        c = ws.cell(11, i, h)
        c.font = font(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center", wrap=True)
        c.border = THIN
    ws.row_dimensions[11].height = 28

    for i in range(NUM_SKILL_SLOTS):
        r = 12 + i
        skill_row = 6 + i
        ws.cell(r, 1, f'=IF({q(SHEET_SKILLS)}!B{skill_row}="","",{q(SHEET_SKILLS)}!B{skill_row})')
        ws.cell(r, 2, f'=IF(A{r}="","",{q(SHEET_SKILLS)}!C{skill_row})')
        ws.cell(r, 3, f'=IF(A{r}="","",IFERROR(ROUND(AVERAGEIF(DataSkill,A{r},DataLastYear),2),""))')
        ws.cell(r, 4, f'=IF(A{r}="","",IFERROR(ROUND(AVERAGEIF(DataSkill,A{r},DataThisYear),2),""))')
        ws.cell(r, 5, f'=IF(OR(C{r}="",D{r}=""),"",ROUND(D{r}-C{r},2))')
        ws.cell(
            r,
            6,
            f'=IF(E{r}="","",IF(E{r}>0,"Improved",IF(E{r}<0,"Declined","Unchanged")))',
        )
        ws.cell(r, 7, f'=IF(A{r}="","",COUNTIFS(DataSkill,A{r},DataThisYear,">=1"))')
        ws.cell(
            r,
            8,
            f'=IF(OR(A{r}="",D{r}=""),"",IF(D{r}<TargetRating,"Yes","No"))',
        )
        for c in range(1, 9):
            cell = ws.cell(r, c)
            cell.border = THIN
            cell.font = font(11)
            cell.alignment = align("center" if c > 2 else "left")
        ws.cell(r, 3).number_format = "0.00"
        ws.cell(r, 4).number_format = "0.00"
        ws.cell(r, 5).number_format = "+0.00;-0.00;0.00"
        ws.row_dimensions[r].height = 20

        ws.conditional_formatting.add(
            f"E{r}",
            CellIsRule(operator="greaterThan", formula=["0"], font=font(11, bold=True, color=GREEN), fill=fill("D4EDDA")),
        )
        ws.conditional_formatting.add(
            f"E{r}",
            CellIsRule(operator="lessThan", formula=["0"], font=font(11, bold=True, color=RED), fill=fill("F8D7DA")),
        )
        ws.conditional_formatting.add(
            f"H{r}",
            CellIsRule(operator="equal", formula=['"Yes"'], font=font(11, bold=True, color=RED), fill=fill("F8D7DA")),
        )
        ws.conditional_formatting.add(
            f"F{r}",
            CellIsRule(operator="equal", formula=['"Improved"'], font=font(11, bold=True, color=GREEN)),
        )
        ws.conditional_formatting.add(
            f"F{r}",
            CellIsRule(operator="equal", formula=['"Declined"'], font=font(11, bold=True, color=RED)),
        )

    # Skill chart — first 10 sample skills (extend range after adding more)
    chart1 = BarChart()
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "Average rating by skill"
    chart1.y_axis.title = "Average (1–5)"
    chart1.y_axis.scaling.min = 0
    chart1.y_axis.scaling.max = 5
    chart1.style = 10
    data = Reference(ws, min_col=3, min_row=11, max_col=4, max_row=11 + len(SKILLS))
    cats = Reference(ws, min_col=1, min_row=12, max_row=11 + len(SKILLS))
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.legend.position = "b"
    chart1.y_axis.majorGridlines = None
    chart1.width = 18
    chart1.height = 8
    if chart1.series:
        chart1.series[0].graphicalProperties.solidFill = "5B8FB9"
        if len(chart1.series) > 1:
            chart1.series[1].graphicalProperties.solidFill = TEAL
    ws.add_chart(chart1, "J10")

    # Employee comparison
    emp_start = 30
    ws.merge_cells(start_row=emp_start, start_column=1, end_row=emp_start, end_column=6)
    ws.cell(emp_start, 1, "Employee comparison — last year vs this year").font = font(14, bold=True, color=NAVY)

    emp_headers = ["Employee", "Department", "Last year avg", "This year avg", "Change", "Below target?"]
    for i, h in enumerate(emp_headers, 1):
        c = ws.cell(emp_start + 1, i, h)
        c.font = font(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center", wrap=True)
        c.border = THIN

    for i in range(NUM_EMPLOYEE_SLOTS):
        r = emp_start + 2 + i
        emp_row = 6 + i
        ws.cell(r, 1, f'=IF({q(SHEET_EMP)}!B{emp_row}="","",{q(SHEET_EMP)}!B{emp_row})')
        ws.cell(r, 2, f'=IF(A{r}="","",{q(SHEET_EMP)}!C{emp_row})')
        ws.cell(r, 3, f'=IF(A{r}="","",IFERROR(ROUND(AVERAGEIF(DataEmployee,A{r},DataLastYear),2),""))')
        ws.cell(r, 4, f'=IF(A{r}="","",IFERROR(ROUND(AVERAGEIF(DataEmployee,A{r},DataThisYear),2),""))')
        ws.cell(r, 5, f'=IF(OR(C{r}="",D{r}=""),"",ROUND(D{r}-C{r},2))')
        ws.cell(
            r,
            6,
            f'=IF(OR(A{r}="",D{r}=""),"",IF(D{r}<TargetRating,"Yes","No"))',
        )
        for c in range(1, 7):
            cell = ws.cell(r, c)
            cell.border = THIN
            cell.font = font(11)
            cell.alignment = align("center" if c > 2 else "left")
        ws.cell(r, 3).number_format = "0.00"
        ws.cell(r, 4).number_format = "0.00"
        ws.cell(r, 5).number_format = "+0.00;-0.00;0.00"
        ws.conditional_formatting.add(
            f"E{r}",
            CellIsRule(operator="greaterThan", formula=["0"], font=font(11, bold=True, color=GREEN), fill=fill("D4EDDA")),
        )
        ws.conditional_formatting.add(
            f"E{r}",
            CellIsRule(operator="lessThan", formula=["0"], font=font(11, bold=True, color=RED), fill=fill("F8D7DA")),
        )
        ws.conditional_formatting.add(
            f"F{r}",
            CellIsRule(operator="equal", formula=['"Yes"'], font=font(11, bold=True, color=RED), fill=fill("F8D7DA")),
        )

    chart2 = BarChart()
    chart2.type = "bar"
    chart2.grouping = "clustered"
    chart2.title = "Average rating by employee"
    chart2.x_axis.title = "Average (1–5)"
    chart2.x_axis.scaling.min = 0
    chart2.x_axis.scaling.max = 5
    chart2.style = 10
    data2 = Reference(ws, min_col=3, min_row=emp_start + 1, max_col=4, max_row=emp_start + 1 + len(EMPLOYEES))
    cats2 = Reference(ws, min_col=1, min_row=emp_start + 2, max_row=emp_start + 1 + len(EMPLOYEES))
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.shape = 4
    chart2.legend.position = "b"
    chart2.width = 18
    chart2.height = 8
    if chart2.series:
        chart2.series[0].graphicalProperties.solidFill = "5B8FB9"
        if len(chart2.series) > 1:
            chart2.series[1].graphicalProperties.solidFill = TEAL
    ws.add_chart(chart2, "J30")

    # Rating distribution
    dist_row = 50
    ws.merge_cells(start_row=dist_row, start_column=1, end_row=dist_row, end_column=4)
    ws.cell(dist_row, 1, "How ratings are distributed").font = font(14, bold=True, color=NAVY)
    for i, h in enumerate(["Rating", "Last year", "This year"], 1):
        c = ws.cell(dist_row + 1, i, h)
        c.font = font(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")
        c.border = THIN
    for score in range(1, 6):
        r = dist_row + 1 + score
        ws.cell(r, 1, score).fill = fill(RATING_FILLS[score])
        ws.cell(r, 1).font = font(12, bold=True)
        ws.cell(r, 1).alignment = align("center")
        ws.cell(r, 2, f"=COUNTIF(DataLastYear,{score})")
        ws.cell(r, 3, f"=COUNTIF(DataThisYear,{score})")
        for c in range(1, 4):
            ws.cell(r, c).border = THIN
            ws.cell(r, c).alignment = align("center")

    chart3 = BarChart()
    chart3.type = "col"
    chart3.grouping = "clustered"
    chart3.title = "Count of ratings 1–5"
    chart3.style = 10
    data3 = Reference(ws, min_col=2, min_row=dist_row + 1, max_col=3, max_row=dist_row + 6)
    cats3 = Reference(ws, min_col=1, min_row=dist_row + 2, max_row=dist_row + 6)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.legend.position = "b"
    chart3.width = 12
    chart3.height = 7
    if chart3.series:
        chart3.series[0].graphicalProperties.solidFill = "5B8FB9"
        if len(chart3.series) > 1:
            chart3.series[1].graphicalProperties.solidFill = TEAL
    ws.add_chart(chart3, "E50")

    note = ws.cell(
        58,
        1,
        "Tip: sort the Skill comparison table by Change or This year avg to see the biggest gaps. "
        "After you add more than 10 skills, right-click a chart → Select Data and extend the range.",
    )
    note.font = font(10, italic=True, color=MUTED)
    ws.merge_cells("A58:H59")
    note.alignment = align("left", wrap=True)

    widths = [24, 16, 16, 16, 12, 14, 12, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col in range(9, 17):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[7].height = 22
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = GOLD
    ws.sheet_view.zoomScale = 100
    apply_print(ws, landscape=True, fit_width=1, fit_height=1)


# ===========================================================================
# How to Use
# ===========================================================================
def box(ws, r1, c1, r2, c2, title, body, accent=TEAL) -> None:
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r1, end_column=c2)
    head = ws.cell(r1, c1, title)
    head.font = font(13, bold=True, color=WHITE)
    head.fill = fill(accent)
    head.alignment = align("left")
    for c in range(c1, c2 + 1):
        ws.cell(r1, c).fill = fill(accent)
    ws.merge_cells(start_row=r1 + 1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(r1 + 1, c1, body)
    cell.font = font(11, color=INK)
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for r in range(r1 + 1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill(PAPER)
            ws.cell(r, c).border = Border(
                left=Side(style="thin", color=LINE),
                right=Side(style="thin", color=LINE),
                top=Side(style="thin", color=LINE) if r == r1 + 1 else Side(style=None),
                bottom=Side(style="thin", color=LINE) if r == r2 else Side(style=None),
            )


def build_how_to(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_HOW, 0)
    last_col = 12
    add_navigation(ws, SHEET_HOW, last_col)

    ws.merge_cells("A2:L2")
    style_title(ws["A2"], "How to use this skill matrix")
    ws.merge_cells("A3:L3")
    ws["A3"] = (
        "Rate each person’s skills from 1 to 5, compare last year with this year, "
        "and add new skills whenever the team grows. Sample people and ratings are included so you can see how it works — replace them with your team."
    )
    ws["A3"].font = font(12, color=MUTED)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 36

    box(
        ws,
        5,
        1,
        14,
        6,
        "1. Rate skills (1–5)",
        "Open the Skill Matrix sheet.\n\n"
        "Each skill has three columns:\n"
        "• Last year — what you recorded last cycle\n"
        "• This year — the current rating\n"
        "• Change — this year minus last year (automatic)\n\n"
        "Type a whole number from 1 to 5. Cells color themselves:\n"
        "1 beginner · 2 basic · 3 competent · 4 proficient · 5 expert\n\n"
        "Green change = improved. Red change = declined.",
        NAVY,
    )
    box(
        ws,
        5,
        7,
        14,
        12,
        "2. Add a skill later",
        "Open the Skills sheet.\n\n"
        "In the next yellow row, enter:\n"
        "• Skill name (required) — this becomes a new column group\n"
        "• Category (optional) — Technical, Tools, Soft Skills, or Domain\n"
        "• Description (optional)\n\n"
        "Go back to Skill Matrix. The next placeholder columns "
        "will show the new skill name.\n\n"
        "Enter last-year and this-year ratings for each employee. "
        "The dashboard updates by itself.\n\n"
        f"{NUM_SKILL_SLOTS} skill slots are ready. If you fill them all, "
        "see “Need more slots?” below.",
        TEAL,
    )
    box(
        ws,
        16,
        1,
        24,
        6,
        "3. Add an employee later",
        "Open the Employees sheet.\n\n"
        "Use the next yellow row. Keep the Employee ID and fill in "
        "Full Name and Department (job title and manager are optional).\n\n"
        "The Skill Matrix adds a row for that person. Rate each skill.\n\n"
        f"{NUM_EMPLOYEE_SLOTS} employee slots are ready.",
        "3D5A80",
    )
    box(
        ws,
        16,
        7,
        24,
        12,
        "4. Compare last year with this year",
        "On Skill Matrix, read the Change column next to each skill.\n\n"
        "Open Dashboard for:\n"
        "• Team average this year vs last year\n"
        "• Each skill’s average and whether it improved\n"
        "• Each person’s average\n"
        "• Skills below the target rating (set on Settings)\n"
        "• How many 1s, 2s, 3s, 4s, and 5s you gave\n\n"
        "Change the years and the target on the Settings sheet.",
        GOLD,
    )

    # Rating scale strip
    ws.merge_cells("A26:L26")
    ws["A26"] = "Rating scale"
    ws["A26"].font = font(14, bold=True, color=NAVY)

    for i, (score, level, meaning) in enumerate(RATING_SCALE):
        c1 = 1 + i * 2
        c2 = c1 + 1
        ws.merge_cells(start_row=27, start_column=c1, end_row=27, end_column=c2)
        ws.merge_cells(start_row=28, start_column=c1, end_row=29, end_column=c2)
        top = ws.cell(27, c1, f"{score}  {level}")
        top.font = font(12, bold=True, color=INK)
        top.fill = fill(RATING_FILLS[score])
        top.alignment = align("center")
        body = ws.cell(28, c1, meaning)
        body.font = font(10, color=INK)
        body.fill = fill(RATING_FILLS[score])
        body.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        for r in range(27, 30):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).fill = fill(RATING_FILLS[score])
                ws.cell(r, c).border = THIN
    ws.row_dimensions[28].height = 20
    ws.row_dimensions[29].height = 28

    ws.merge_cells("A31:L33")
    ws["A31"] = (
        "Need more slots? On Skills or Employees, click a row inside the table and use "
        "Home → Insert → Insert Table Rows Below. Then copy the last Skill Matrix column group "
        "(three columns: last year, this year, change) and insert the copy to the right. "
        "Point the new header formulas at the extra Skills row, and copy the Change formula down. "
        "Most teams will not need this — 16 skills and 16 people are already set up.\n\n"
        "A hidden Data sheet powers the dashboard. You do not need to edit it."
    )
    ws["A31"].font = font(11, color=MUTED)
    ws["A31"].alignment = align("left", wrap=True)

    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.row_dimensions[2].height = 28
    for r in (6, 17):
        ws.row_dimensions[r].height = 20
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TEAL
    ws.freeze_panes = "A2"
    apply_print(ws, landscape=True, fit_width=1, fit_height=1)


def main() -> None:
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    build_how_to(wb)
    build_matrix(wb)
    build_dashboard(wb)
    build_employees(wb)
    build_skills(wb)
    build_settings(wb)
    build_data(wb)

    # Keep a friendly sheet order
    order = [SHEET_HOW, SHEET_MATRIX, SHEET_DASH, SHEET_EMP, SHEET_SKILLS, SHEET_SETTINGS, SHEET_DATA]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.properties.title = "Employee Skill Matrix"
    wb.properties.creator = "Skill Matrix"
    wb.properties.description = (
        "Rate employee skills from 1–5, add skills over time, and compare last year with this year."
    )

    path = "Employee_Skill_Matrix.xlsx"
    wb.save(path)
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
