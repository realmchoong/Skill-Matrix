#!/usr/bin/env python3
"""Generate a ready-to-use Employee Skill Matrix Excel workbook."""

from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------
NUM_SKILL_SLOTS = 16
NUM_EMPLOYEE_SLOTS = 16
NUM_YEAR_SLOTS = 16
NUM_LIST_SLOTS = 16
NUM_SET_SLOTS = 40
MAX_SET_SKILLS = 6

NAV_ROW = 1
TITLE_ROW = 2
MATRIX_INFO_ROW = 3
MATRIX_HEADER_ROW = 5
MATRIX_DATA_START = 6
MATRIX_ROWS = NUM_EMPLOYEE_SLOTS * NUM_SKILL_SLOTS
MATRIX_DATA_END = MATRIX_DATA_START + MATRIX_ROWS - 1
YEAR_FIRST_COL = 5  # E
SUMMARY_FIRST_COL = YEAR_FIRST_COL + NUM_YEAR_SLOTS

SHEET_HOW = "How to Use"
SHEET_MATRIX = "Skill Matrix"
SHEET_RATINGS = "Ratings"
SHEET_DASH = "Dashboard"
SHEET_EMP = "Employees"
SHEET_SKILLS = "Skills"
SHEET_YEARS = "Years"
SHEET_LISTS = "Lists"
SHEET_SETS = "Skill Sets"
SHEET_SETTINGS = "Settings"
SHEET_DATA = "Data"
SHEET_CALC = "Calc"

NAV_SHEETS = [
    SHEET_HOW,
    SHEET_MATRIX,
    SHEET_RATINGS,
    SHEET_DASH,
    SHEET_EMP,
    SHEET_SKILLS,
    SHEET_YEARS,
    SHEET_LISTS,
    SHEET_SETS,
    SHEET_SETTINGS,
]

GLANCE_PAIR_ROW = 3
GLANCE_CAT_ROW = 4
GLANCE_SKILL_ROW = 5
GLANCE_SUB_ROW = 6
GLANCE_START = 7
GLANCE_END = GLANCE_START + NUM_EMPLOYEE_SLOTS - 1

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
NAVY = "1B365D"
TEAL = "0E7C7B"
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
YELLOW = "FFF8E7"

RATING_FILLS = {
    1: "F5D0D0",
    2: "FAD9C4",
    3: "FBF3D0",
    4: "D4EDDA",
    5: "A8D5BA",
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
TARGET = 3
PRESET_YEARS = list(range(2023, 2035))  # already waiting as columns

# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------
EMPLOYEES = [
    ("EMP-001", "Alex Rivera", "Vision", "Vision Switcher", "Riley Thompson", date(2022, 3, 12)),
    ("EMP-002", "Jordan Chen", "Vision", "Camera Switcher", "Riley Thompson", date(2021, 6, 15)),
    ("EMP-003", "Sam Patel", "Content", "Content Operator", "Riley Thompson", date(2023, 1, 10)),
    ("EMP-004", "Taylor Brooks", "Systems", "Systems Technician", "Riley Thompson", date(2020, 9, 1)),
    ("EMP-005", "Morgan Lee", "Systems", "Setup Technician", "Riley Thompson", date(2024, 2, 20)),
    ("EMP-006", "Casey Nguyen", "Streaming", "Streaming Operator", "Riley Thompson", date(2022, 11, 4)),
    ("EMP-007", "Riley Thompson", "Vision", "Vision Supervisor", "Dana Wright", date(2019, 4, 8)),
    ("EMP-008", "Avery Kim", "Content", "Content Operator", "Riley Thompson", date(2024, 8, 18)),
]

SKILLS = [
    ("SK-001", "Barco Vision Switching", "Switching", "Operate Barco (or similar) processors and switch LED / IMAG / program during the show"),
    ("SK-002", "Camera Switching", "Switching", "Cut cameras and direct the live picture"),
    ("SK-003", "Content Operation", "Content", "Cue and play out show content — media servers, playback, and graphics"),
    ("SK-004", "Systems Tech", "Systems", "Signal flow, routing, networking, and show systems"),
    ("SK-005", "Technical Setup", "Systems", "Rig, patch, and line-check vision systems before doors"),
    ("SK-006", "Live Streaming", "Streaming", "Encode, monitor, and deliver the live stream"),
    ("SK-007", "Troubleshooting", "Support", "Diagnose and fix issues under show pressure"),
]

DEPARTMENTS = ["Vision", "Content", "Systems", "Streaming"]
JOB_TITLES = [
    "Vision Switcher",
    "Camera Switcher",
    "Content Operator",
    "Systems Technician",
    "Setup Technician",
    "Streaming Operator",
    "Vision Supervisor",
]

# (2025, 2026) for: Barco, Camera, Content, Systems, Setup, Streaming, Troubleshooting
BASE_2526 = [
    [(4, 5), (3, 3), (2, 3), (3, 3), (3, 4), (2, 2), (3, 4)],
    [(3, 3), (5, 5), (2, 3), (2, 3), (3, 3), (2, 2), (4, 4)],
    [(2, 2), (2, 3), (4, 5), (2, 3), (3, 3), (4, 4), (3, 3)],
    [(2, 3), (2, 2), (3, 3), (5, 5), (4, 5), (3, 3), (4, 5)],
    [(1, 2), (1, 2), (2, 3), (3, 4), (4, 5), (2, 2), (3, 4)],
    [(1, 2), (2, 2), (3, 4), (3, 3), (3, 3), (4, 5), (3, 4)],
    [(4, 4), (4, 4), (3, 4), (4, 4), (4, 4), (3, 3), (5, 5)],
    [(1, 2), (2, 2), (4, 5), (2, 3), (2, 3), (3, 4), (3, 3)],
]

SKILL_SETS = [
    (
        "Barco Vision Switching",
        "Lead Barco / LED / program switching on a live event",
        [("Barco Vision Switching", 4), ("Technical Setup", 3), ("Troubleshooting", 3)],
    ),
    (
        "Camera Switching",
        "Cut cameras and own the live picture",
        [("Camera Switching", 4), ("Technical Setup", 3), ("Troubleshooting", 3)],
    ),
    (
        "Content Operation",
        "Cue and play out show content",
        [("Content Operation", 4), ("Live Streaming", 3), ("Troubleshooting", 3)],
    ),
    (
        "Systems Tech",
        "Own signal flow, routing, and show systems",
        [("Systems Tech", 4), ("Technical Setup", 4), ("Troubleshooting", 4)],
    ),
    (
        "Technical Setup",
        "Rig, patch, and line-check before doors",
        [("Technical Setup", 4), ("Systems Tech", 3), ("Troubleshooting", 3)],
    ),
    (
        "Live Streaming",
        "Encode, monitor, and deliver the live stream",
        [("Live Streaming", 4), ("Content Operation", 3), ("Troubleshooting", 3)],
    ),
    (
        "Troubleshooting",
        "Diagnose and fix issues under show pressure",
        [("Troubleshooting", 4), ("Systems Tech", 3), ("Technical Setup", 3)],
    ),
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
    return Font(name=FONT, size=size, bold=bold, italic=italic, color=color, underline=underline)


def align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def q(sheet: str) -> str:
    return f"'{sheet}'"


def glance_skill_columns(n: int) -> tuple[int, int, int]:
    """1-based (earlier, later, change) columns for skill slot n (1-based)."""
    start = 3 + (n - 1) * 3
    return start, start + 1, start + 2


def year_col(slot: int) -> int:
    return YEAR_FIRST_COL + slot


def matrix_row(emp_i: int, skill_i: int) -> int:
    return MATRIX_DATA_START + emp_i * NUM_SKILL_SLOTS + skill_i


def add_defined_name(wb: Workbook, name: str, ref: str) -> None:
    wb.defined_names.add(DefinedName(name=name, attr_text=ref))


def add_navigation(ws, active: str, last_col: int) -> None:
    for i, name in enumerate(NAV_SHEETS, start=1):
        cell = ws.cell(NAV_ROW, i, name)
        cell.hyperlink = f"#'{name}'!A1"
        cell.font = font(10, bold=name == active, color=WHITE, underline="single")
        cell.fill = fill(TEAL if name == active else NAVY)
        cell.alignment = align("center")
        cell.border = Border(
            left=Side(style="thin", color="0A3A5C"),
            right=Side(style="thin", color="0A3A5C"),
        )
    for col in range(len(NAV_SHEETS) + 1, last_col + 1):
        ws.cell(NAV_ROW, col).fill = fill(NAVY)
    ws.row_dimensions[NAV_ROW].height = 24


def style_title(cell, text: str, size: int = 22) -> None:
    cell.value = text
    cell.font = font(size, bold=True, color=NAVY)
    cell.alignment = align("left", "center")


def apply_print(ws, landscape=True) -> None:
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
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


def rating_for(emp_i: int, skill_i: int, year: int):
    if emp_i >= len(BASE_2526) or skill_i >= len(BASE_2526[0]):
        return None
    hire_year = EMPLOYEES[emp_i][5].year
    if year < hire_year:
        return None
    y25, y26 = BASE_2526[emp_i][skill_i]
    if year == 2026:
        return y26
    if year == 2025:
        return y25
    if year == 2024:
        return max(1, y25 - (1 if (emp_i + skill_i) % 3 == 0 else 0))
    if year == 2023:
        r24 = max(1, y25 - (1 if (emp_i + skill_i) % 3 == 0 else 0))
        return max(1, r24 - (1 if (emp_i + skill_i) % 2 == 0 else 0))
    return None


def style_input(cell, value=None) -> None:
    if value is not None:
        cell.value = value
    cell.fill = fill(YELLOW)
    cell.border = MED
    cell.alignment = align("center")
    cell.font = font(12, bold=True, color=TEAL)


def add_rating_validation(ws, cells_range: str) -> None:
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
    dv.error = "Enter a whole number from 1 (beginner) to 5 (expert), or leave blank."
    dv.promptTitle = "Skill rating"
    dv.prompt = "1 beginner · 2 basic · 3 competent · 4 proficient · 5 expert"
    ws.add_data_validation(dv)
    dv.add(cells_range)


def apply_rating_cf(ws, cells_range: str) -> None:
    for score, color in RATING_FILLS.items():
        ws.conditional_formatting.add(
            cells_range,
            CellIsRule(
                operator="equal",
                formula=[str(score)],
                fill=fill(color),
                font=font(12, bold=True, color=INK),
            ),
        )


def apply_change_cf(ws, cells_range: str) -> None:
    ws.conditional_formatting.add(
        cells_range,
        CellIsRule(operator="greaterThan", formula=["0"], fill=fill("D4EDDA"), font=font(11, bold=True, color=GREEN)),
    )
    ws.conditional_formatting.add(
        cells_range,
        CellIsRule(operator="lessThan", formula=["0"], fill=fill("F8D7DA"), font=font(11, bold=True, color=RED)),
    )
    ws.conditional_formatting.add(
        cells_range,
        CellIsRule(operator="equal", formula=["0"], fill=fill("E9ECEF"), font=font(11, color=MUTED)),
    )


# ===========================================================================
# Settings
# ===========================================================================
def build_settings(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_SETTINGS)
    add_navigation(ws, SHEET_SETTINGS, 8)
    ws.merge_cells("A2:H2")
    style_title(ws["A2"], "Settings")
    ws.merge_cells("A3:H3")
    ws["A3"] = "Target rating is the only number you usually need to change. Years are managed on the Years sheet."
    ws["A3"].font = font(11, color=MUTED, italic=True)

    headers = ["Setting", "Value", "What it controls"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")

    ws["A6"] = "Target rating"
    ws["B6"] = TARGET
    ws["C6"] = "Latest scores below this are flagged as development needs on the Dashboard."
    style_input(ws["B6"], TARGET)
    ws["A6"].font = font(11, bold=True)
    ws["A6"].fill = fill(PAPER)
    ws["A6"].border = THIN
    ws["C6"].font = font(11, color=MUTED)
    ws.merge_cells("C6:H6")

    ws["A8"] = "Read-only (calculated)"
    ws["A8"].font = font(14, bold=True, color=NAVY)
    ws.merge_cells("A8:C8")

    ws["A9"] = "Calendar year"
    ws["B9"] = "=YEAR(TODAY())"
    ws["C9"] = "Today's calendar year. Used to highlight the current column and prompt for a new year."
    ws["A10"] = "Latest year with ratings"
    ws["B10"] = "=LatestYear"
    ws["C10"] = "Most recent year that has at least one rating entered."
    ws["A11"] = "Previous year with ratings"
    ws["B11"] = "=PreviousYear"
    ws["C11"] = "The year before that — used for last-year comparison."
    for r in range(9, 12):
        ws.cell(r, 1).font = font(11, bold=True)
        ws.cell(r, 1).fill = fill(PAPER)
        ws.cell(r, 1).border = THIN
        ws.cell(r, 2).font = font(14, bold=True, color=TEAL)
        ws.cell(r, 2).alignment = align("center")
        ws.cell(r, 2).border = THIN
        ws.cell(r, 3).font = font(11, color=MUTED)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)

    ws["A13"] = "Rating scale (1–5)"
    ws["A13"].font = font(14, bold=True, color=NAVY)
    for i, h in enumerate(["Rating", "Level", "Meaning"], 1):
        c = ws.cell(14, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")
    for i, (score, level, meaning) in enumerate(RATING_SCALE):
        r = 15 + i
        for col, val in ((1, score), (2, level), (3, meaning)):
            c = ws.cell(r, col, val)
            c.fill = fill(RATING_FILLS[score])
            c.border = THIN
            c.font = font(16 if col == 1 else 11, bold=col < 3)
            c.alignment = align("center" if col < 3 else "left", wrap=True)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 28

    add_defined_name(wb, "TargetRating", f"{q(SHEET_SETTINGS)}!$B$6")
    add_defined_name(wb, "CalendarYear", f"{q(SHEET_SETTINGS)}!$B$9")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    for col in range(3, 9):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.freeze_panes = "A2"
    ws.sheet_properties.tabColor = GOLD
    apply_print(ws, landscape=False)


# ===========================================================================
# Lists (departments & job titles)
# ===========================================================================
def build_lists(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_LISTS)
    add_navigation(ws, SHEET_LISTS, 8)
    ws.merge_cells("A2:H2")
    style_title(ws["A2"], "Dropdown lists")
    ws.merge_cells("A3:H3")
    ws["A3"] = (
        "Add a department or job title in the next yellow row. "
        "It appears in the Employees dropdowns automatically. "
        "Add it here first, then choose it on the Employees sheet."
    )
    ws["A3"].font = font(11, color=MUTED, italic=True)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 36

    ws["A5"] = "Department"
    ws["C5"] = "Job title"
    ws["A5"].font = font(11, bold=True, color=WHITE)
    ws["C5"].font = font(11, bold=True, color=WHITE)
    ws["A5"].fill = fill(NAVY)
    ws["C5"].fill = fill(NAVY)
    ws["A5"].alignment = align("center")
    ws["C5"].alignment = align("center")

    for i in range(NUM_LIST_SLOTS):
        r = 6 + i
        dept = DEPARTMENTS[i] if i < len(DEPARTMENTS) else ""
        title = JOB_TITLES[i] if i < len(JOB_TITLES) else ""
        dcell = ws.cell(r, 1, dept)
        tcell = ws.cell(r, 3, title)
        for cell in (dcell, tcell):
            cell.font = font(11)
            cell.border = THIN
            if not cell.value:
                cell.fill = fill(YELLOW)
        ws.cell(r, 2).border = Border()

    ws.cell(6 + len(DEPARTMENTS), 1).comment = Comment(
        "Type a new department here. It will appear in the Department dropdown on Employees.",
        "Skill Matrix",
        width=260,
        height=70,
    )
    ws.cell(6 + len(JOB_TITLES), 3).comment = Comment(
        "Type a new job title here. It will appear in the Job Title dropdown on Employees.",
        "Skill Matrix",
        width=260,
        height=70,
    )

    add_table(ws, "tblDepartments", f"A5:A{5 + NUM_LIST_SLOTS}")
    add_table(ws, "tblJobTitles", f"C5:C{5 + NUM_LIST_SLOTS}")

    # Filter lists that start with (All ...)
    ws["E5"] = "DeptFilter"
    ws["E6"] = "(All departments)"
    for i in range(NUM_LIST_SLOTS):
        ws.cell(7 + i, 5, f'=IF(A{6 + i}="","",A{6 + i})')
    ws.column_dimensions["E"].hidden = True

    last_dept = 5 + NUM_LIST_SLOTS
    last_title = 5 + NUM_LIST_SLOTS
    add_defined_name(
        wb,
        "DepartmentList",
        f"{q(SHEET_LISTS)}!$A$6:INDEX({q(SHEET_LISTS)}!$A$6:$A${last_dept},COUNTA({q(SHEET_LISTS)}!$A$6:$A${last_dept}))",
    )
    add_defined_name(
        wb,
        "JobTitleList",
        f"{q(SHEET_LISTS)}!$C$6:INDEX({q(SHEET_LISTS)}!$C$6:$C${last_title},COUNTA({q(SHEET_LISTS)}!$C$6:$C${last_title}))",
    )
    add_defined_name(
        wb,
        "DeptFilterList",
        f"{q(SHEET_LISTS)}!$E$6:INDEX({q(SHEET_LISTS)}!$E$6:$E${6 + NUM_LIST_SLOTS},COUNTA({q(SHEET_LISTS)}!$E$6:$E${6 + NUM_LIST_SLOTS}))",
    )

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 4
    ws.column_dimensions["C"].width = 28
    ws.freeze_panes = "A5"
    ws.sheet_properties.tabColor = "3D5A80"
    apply_print(ws, landscape=False)


# ===========================================================================
# Years
# ===========================================================================
def build_years(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_YEARS)
    add_navigation(ws, SHEET_YEARS, 8)
    ws.merge_cells("A2:D2")
    style_title(ws["A2"], "Years")

    # Call-to-action panel that behaves like an Add Year button
    ws.merge_cells("F2:H4")
    btn = ws["F2"]
    btn.value = '=IF(COUNTIF(B:B,YEAR(TODAY()))>0,"CURRENT YEAR IS READY","ADD THIS YEAR")'
    btn.font = font(16, bold=True, color=WHITE)
    btn.fill = fill(TEAL)
    btn.alignment = align("center", wrap=True)
    for r in range(2, 5):
        for c in range(6, 9):
            ws.cell(r, c).fill = fill(TEAL)
            ws.cell(r, c).border = MED
            ws.cell(r, c).font = font(16, bold=True, color=WHITE)
            ws.cell(r, c).alignment = align("center", wrap=True)

    ws.merge_cells("A3:D4")
    ws["A3"] = (
        '=IF(COUNTIF($B$8:$B$23,YEAR(TODAY()))>0,'
        '"Calendar year "&YEAR(TODAY())&" is already a column on Ratings. '
        'Open Ratings and enter this year\'s 1–5 scores. Then on Skill Matrix pick a year pair '
        '(for example "&YEAR(TODAY())-1&" vs "&YEAR(TODAY())&") to glance at the whole crew. '
        'Next years through 2034 are already waiting.",'
        '"New calendar year "&YEAR(TODAY())&" is not in the list yet. Type "&YEAR(TODAY())'
        '&" in the next yellow row in the Year column, then rate people on the Skill Matrix.")'
    )
    ws["A3"].font = font(11, color=INK)
    ws["A3"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws["A3"].fill = fill(PAPER)

    ws.merge_cells("F5:H6")
    ws["F5"] = (
        '=IF(COUNTIF($B$8:$B$23,YEAR(TODAY()))>0,'
        'HYPERLINK("#\'Ratings\'!A1","Rate "&YEAR(TODAY())&" on Ratings →"),'
        'HYPERLINK("#Years!B20","Type "&YEAR(TODAY())&" in the yellow row →"))'
    )
    ws["F5"].font = font(13, bold=True, color=WHITE, underline="single")
    ws["F5"].fill = fill(NAVY)
    ws["F5"].alignment = align("center")
    for r in range(5, 7):
        for c in range(6, 9):
            ws.cell(r, c).fill = fill(NAVY)

    ws.merge_cells("A5:D6")
    ws["A5"] = (
        "To add a year beyond 2034 (or a missing historical year): type it in the next yellow row, "
        "in order, oldest to newest. A new column appears on Ratings, and a new pair on Skill Matrix. "
        "Do not use a formula for a year that already has ratings."
    )
    ws["A5"].font = font(10, italic=True, color=MUTED)
    ws["A5"].alignment = align("left", wrap=True)

    ws["A7"] = "#"
    ws["B7"] = "Year"
    ws["C7"] = "Label"
    ws["D7"] = "Notes"
    for col in range(1, 5):
        c = ws.cell(7, col)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")

    labels = {2026: "Current sample year", 2025: "Last sample year", 2024: "Two years back", 2023: "Three years back"}
    for i in range(NUM_YEAR_SLOTS):
        r = 8 + i
        ws.cell(r, 1, i + 1).alignment = align("center")
        ws.cell(r, 1).border = THIN
        if i < len(PRESET_YEARS):
            year = PRESET_YEARS[i]
            cell = ws.cell(r, 2, year)
            cell.font = font(12, bold=True)
            ws.cell(r, 3, labels.get(year, "Ready for ratings"))
            ws.cell(r, 4, "Column already on Skill Matrix")
        else:
            cell = ws.cell(r, 2, "")
            cell.fill = fill(YELLOW)
            ws.cell(r, 3, "")
            ws.cell(r, 4, "Type a year here to add a column")
            ws.cell(r, 3).fill = fill(YELLOW)
            ws.cell(r, 4).fill = fill(YELLOW)
        for c in range(2, 5):
            ws.cell(r, c).border = THIN
            ws.cell(r, c).font = font(11)
        cell.alignment = align("center")
        cell.number_format = "0"

    first_empty = 8 + len(PRESET_YEARS)
    ws.cell(first_empty, 2).comment = Comment(
        "Type a new year here (for example 2035). Keep years in order, oldest at the top. "
        "The Skill Matrix and Dashboard pick it up automatically.",
        "Skill Matrix",
        width=280,
        height=90,
    )

    add_table(ws, "tblYears", f"A7:D{7 + NUM_YEAR_SLOTS}")
    add_defined_name(
        wb,
        "YearList",
        f"{q(SHEET_YEARS)}!$B$8:INDEX({q(SHEET_YEARS)}!$B$8:$B${7 + NUM_YEAR_SLOTS},COUNTA({q(SHEET_YEARS)}!$B$8:$B${7 + NUM_YEAR_SLOTS}))",
    )

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 3
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 16
    ws.freeze_panes = "A8"
    ws.sheet_properties.tabColor = TEAL
    apply_print(ws, landscape=False)


# ===========================================================================
# Employees
# ===========================================================================
def build_employees(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_EMP)
    add_navigation(ws, SHEET_EMP, 7)
    ws.merge_cells("A2:G2")
    style_title(ws["A2"], "Employees")
    ws.merge_cells("A3:G3")
    ws["A3"] = (
        "Add a person in the next yellow row. Department and Job title must be chosen from the dropdowns "
        "(add new options on the Lists sheet first). Hire date is used for “since they joined” on the Dashboard."
    )
    ws["A3"].font = font(11, color=MUTED, italic=True)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 36

    headers = ["Employee ID", "Full Name", "Department", "Job Title", "Manager", "Hire Date", "Status"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")

    for i in range(NUM_EMPLOYEE_SLOTS):
        r = 6 + i
        emp_id = f"EMP-{i + 1:03d}"
        if i < len(EMPLOYEES):
            _, name, dept, title, manager, hired = EMPLOYEES[i]
            status = "Active"
        else:
            name = dept = title = manager = status = ""
            hired = None
        values = [emp_id, name, dept, title, manager, hired, status]
        for c, val in enumerate(values, 1):
            cell = ws.cell(r, c, val)
            cell.font = font(11)
            cell.border = THIN
            if i >= len(EMPLOYEES):
                cell.fill = fill(YELLOW)
        ws.cell(r, 6).number_format = "YYYY-MM-DD"
        ws.cell(r, 1).alignment = align("center")

    ws.cell(6 + len(EMPLOYEES), 2).comment = Comment(
        "Type a new employee's name here, then pick Department and Job title from the dropdowns.",
        "Skill Matrix",
        width=260,
        height=70,
    )

    dept_dv = DataValidation(type="list", formula1="=DepartmentList", allow_blank=True, showErrorMessage=True)
    dept_dv.errorTitle = "Department not in list"
    dept_dv.error = "Add this department on the Lists sheet first, then choose it here."
    dept_dv.promptTitle = "Department"
    dept_dv.prompt = "Choose from the list. Add new departments on the Lists sheet."
    ws.add_data_validation(dept_dv)
    dept_dv.add(f"C6:C{5 + NUM_EMPLOYEE_SLOTS}")

    title_dv = DataValidation(type="list", formula1="=JobTitleList", allow_blank=True, showErrorMessage=True)
    title_dv.errorTitle = "Job title not in list"
    title_dv.error = "Add this job title on the Lists sheet first, then choose it here."
    title_dv.promptTitle = "Job title"
    title_dv.prompt = "Choose from the list. Add new titles on the Lists sheet."
    ws.add_data_validation(title_dv)
    title_dv.add(f"D6:D{5 + NUM_EMPLOYEE_SLOTS}")

    add_table(ws, "tblEmployees", f"A5:G{5 + NUM_EMPLOYEE_SLOTS}")
    add_defined_name(
        wb,
        "EmployeeList",
        f"{q(SHEET_EMP)}!$B$6:INDEX({q(SHEET_EMP)}!$B$6:$B${5 + NUM_EMPLOYEE_SLOTS},COUNTA({q(SHEET_EMP)}!$B$6:$B${5 + NUM_EMPLOYEE_SLOTS}))",
    )

    widths = [14, 22, 16, 22, 18, 14, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = "3D5A80"
    apply_print(ws, landscape=True)


# ===========================================================================
# Skills
# ===========================================================================
def build_skills(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_SKILLS)
    add_navigation(ws, SHEET_SKILLS, 4)
    ws.merge_cells("A2:D2")
    style_title(ws["A2"], "Skills catalog")
    ws.merge_cells("A3:D3")
    ws["A3"] = (
        "Add a skill in the next yellow row. Every employee gets a row for it on the Skill Matrix. "
        "Then add it to a skill set on the Skill Sets sheet if you want to find who can do that job."
    )
    ws["A3"].font = font(11, color=MUTED, italic=True)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 40

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
        for c, val in enumerate([skill_id, name, category, desc], 1):
            cell = ws.cell(r, c, val)
            cell.font = font(11)
            cell.border = THIN
            cell.alignment = align("left" if c > 1 else "center", wrap=True)
            if i >= len(SKILLS):
                cell.fill = fill(YELLOW)
        ws.row_dimensions[r].height = 22

    ws.cell(6 + len(SKILLS), 2).comment = Comment(
        "Type a new skill name here. It appears on the Skill Matrix for every employee.",
        "Skill Matrix",
        width=260,
        height=70,
    )

    cat_dv = DataValidation(type="list", formula1='"Switching,Content,Systems,Streaming,Support"', allow_blank=True)
    cat_dv.prompt = "Optional grouping"
    cat_dv.promptTitle = "Category"
    ws.add_data_validation(cat_dv)
    cat_dv.add(f"C6:C{5 + NUM_SKILL_SLOTS}")

    add_table(ws, "tblSkills", f"A5:D{5 + NUM_SKILL_SLOTS}")
    add_defined_name(
        wb,
        "SkillNameList",
        f"{q(SHEET_SKILLS)}!$B$6:INDEX({q(SHEET_SKILLS)}!$B$6:$B${5 + NUM_SKILL_SLOTS},COUNTA({q(SHEET_SKILLS)}!$B$6:$B${5 + NUM_SKILL_SLOTS}))",
    )

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 62
    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = TEAL
    apply_print(ws, landscape=False)


# ===========================================================================
# Skill Sets
# ===========================================================================
def build_skill_sets(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_SETS)
    add_navigation(ws, SHEET_SETS, 8)
    ws.merge_cells("A2:H2")
    style_title(ws["A2"], "Skill sets (job profiles)")
    ws.merge_cells("A3:H3")
    ws["A3"] = (
        "A skill set is the mix of skills needed for a job. Add rows with the same set name and one skill per row. "
        "On the Dashboard, pick a set to see who currently rates highest."
    )
    ws["A3"].font = font(11, color=MUTED, italic=True)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 40

    ws["A5"] = "Skill set"
    ws["B5"] = "Skill"
    ws["C5"] = "Minimum rating"
    for col in range(1, 4):
        c = ws.cell(5, col)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center")

    row = 6
    for set_name, _desc, skills in SKILL_SETS:
        for skill, minimum in skills:
            ws.cell(row, 1, set_name).border = THIN
            ws.cell(row, 2, skill).border = THIN
            cell = ws.cell(row, 3, minimum)
            cell.border = THIN
            cell.alignment = align("center")
            row += 1
    while row < 6 + NUM_SET_SLOTS:
        for c in range(1, 4):
            cell = ws.cell(row, c, "")
            cell.border = THIN
            cell.fill = fill(YELLOW)
        row += 1

    first_empty = 6 + sum(len(s[2]) for s in SKILL_SETS)
    ws.cell(first_empty, 1).comment = Comment(
        "Add another skill to a set (reuse the set name) or start a new set name here.",
        "Skill Matrix",
        width=260,
        height=70,
    )

    skill_dv = DataValidation(type="list", formula1="=SkillNameList", allow_blank=True)
    skill_dv.prompt = "Choose a skill from the Skills catalog"
    skill_dv.promptTitle = "Skill"
    ws.add_data_validation(skill_dv)
    skill_dv.add(f"B6:B{5 + NUM_SET_SLOTS}")

    min_dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True)
    ws.add_data_validation(min_dv)
    min_dv.add(f"C6:C{5 + NUM_SET_SLOTS}")

    add_table(ws, "tblSkillSetSkills", f"A5:C{5 + NUM_SET_SLOTS}")

    # Unique set names + descriptions for the dashboard dropdown
    ws["E5"] = "Skill set name"
    ws["F5"] = "Description"
    for col in range(5, 7):
        c = ws.cell(5, col)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)
    for i, (name, desc, _skills) in enumerate(SKILL_SETS):
        ws.cell(6 + i, 5, name).border = THIN
        ws.cell(6 + i, 6, desc).border = THIN
    name_slots = 16
    for i in range(len(SKILL_SETS), name_slots):
        ws.cell(6 + i, 5, "").fill = fill(YELLOW)
        ws.cell(6 + i, 6, "").fill = fill(YELLOW)
        ws.cell(6 + i, 5).border = THIN
        ws.cell(6 + i, 6).border = THIN
    add_table(ws, "tblSkillSetNames", f"E5:F{5 + name_slots}")
    add_defined_name(
        wb,
        "SkillSetList",
        f"{q(SHEET_SETS)}!$E$6:INDEX({q(SHEET_SETS)}!$E$6:$E${5 + name_slots},COUNTA({q(SHEET_SETS)}!$E$6:$E${5 + name_slots}))",
    )

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 40
    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = GOLD
    apply_print(ws, landscape=True)


# ===========================================================================
# Skill Matrix (two-year glance of the whole crew)
# ===========================================================================
def build_matrix(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_MATRIX, 0)
    last_col = 2 + NUM_SKILL_SLOTS * 3
    add_navigation(ws, SHEET_MATRIX, last_col)

    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=8)
    style_title(ws["A2"], "Crew skill matrix")
    ws.merge_cells(start_row=TITLE_ROW, start_column=9, end_row=TITLE_ROW, end_column=last_col)
    ws.cell(
        TITLE_ROW,
        9,
        "Every person in one view. Only two years show at a time. Change the Year pair dropdown "
        "to scroll every employee together. Type ratings on the Ratings sheet. No macros.",
    ).font = font(10, italic=True, color=MUTED)
    ws.row_dimensions[TITLE_ROW].height = 28

    # Year-pair scroller (dropdown — works without VBA)
    ws["A3"] = "Scroll years ▼"
    ws["A3"].font = font(11, bold=True, color=WHITE)
    ws["A3"].fill = fill(TEAL)
    ws["A3"].alignment = align("center")
    ws.merge_cells("B3:C3")
    pair = ws["B3"]
    pair.value = "2025 vs 2026"
    style_input(pair)
    pair.font = font(16, bold=True, color=NAVY)
    ws["C3"].fill = fill(YELLOW)
    ws["C3"].border = MED

    ws["D3"] = "Earlier"
    ws["E3"] = '=IFERROR(--LEFT($B$3,4),"")'
    ws["F3"] = "Later"
    ws["G3"] = '=IFERROR(--RIGHT($B$3,4),"")'
    ws["D3"].font = font(10, bold=True, color=WHITE)
    ws["D3"].fill = fill("5B8FB9")
    ws["D3"].alignment = align("center")
    ws["F3"].font = font(10, bold=True, color=WHITE)
    ws["F3"].fill = fill(TEAL)
    ws["F3"].alignment = align("center")
    ws["E3"].font = font(18, bold=True, color=NAVY)
    ws["E3"].fill = fill(LY_FILL)
    ws["E3"].alignment = align("center")
    ws["E3"].number_format = "0"
    ws["E3"].border = MED
    ws["G3"].font = font(18, bold=True, color=NAVY)
    ws["G3"].fill = fill(TY_FILL)
    ws["G3"].alignment = align("center")
    ws["G3"].number_format = "0"
    ws["G3"].border = MED

    ws.merge_cells("H3:L3")
    ws["H3"] = (
        '=HYPERLINK("#\'Ratings\'!A1","Type 1–5 on Ratings →")'
        '&"     Next pair: "&IFERROR(INDEX(YearPairList,MATCH($B$3,YearPairList,0)+1),"—")'
        '&"     Previous pair: "&IFERROR(INDEX(YearPairList,MATCH($B$3,YearPairList,0)-1),"—")'
    )
    ws["H3"].font = font(11, color=MUTED)
    ws["H3"].alignment = align("left")

    add_defined_name(wb, "YearPair", f"{q(SHEET_MATRIX)}!$B$3")
    add_defined_name(wb, "ViewEarlierYear", f"{q(SHEET_MATRIX)}!$E$3")
    add_defined_name(wb, "ViewLaterYear", f"{q(SHEET_MATRIX)}!$G$3")

    # Employee / department headers
    ws.merge_cells(start_row=GLANCE_CAT_ROW, start_column=1, end_row=GLANCE_SUB_ROW, end_column=1)
    ws.merge_cells(start_row=GLANCE_CAT_ROW, start_column=2, end_row=GLANCE_SUB_ROW, end_column=2)
    for col, label in ((1, "Employee"), (2, "Department")):
        cell = ws.cell(GLANCE_CAT_ROW, col, label)
        cell.font = font(12, bold=True, color=WHITE)
        cell.fill = fill(NAVY)
        cell.alignment = align("center", wrap=True)
        cell.border = THIN
        for r in (GLANCE_SKILL_ROW, GLANCE_SUB_ROW):
            ws.cell(r, col).fill = fill(NAVY)
            ws.cell(r, col).border = THIN

    ratings_end = MATRIX_DATA_END
    emp_rng = f"{q(SHEET_RATINGS)}!$A$6:$A${ratings_end}"
    skill_rng = f"{q(SHEET_RATINGS)}!$C$6:$C${ratings_end}"
    year_hdr = f"{q(SHEET_RATINGS)}!$E$5:$T$5"
    year_block = f"{q(SHEET_RATINGS)}!$E$6:$T${ratings_end}"

    for n in range(1, NUM_SKILL_SLOTS + 1):
        earlier_c, later_c, change_c = glance_skill_columns(n)
        skills_row = 5 + n
        ws.merge_cells(
            start_row=GLANCE_CAT_ROW, start_column=earlier_c, end_row=GLANCE_CAT_ROW, end_column=change_c
        )
        cat = ws.cell(
            GLANCE_CAT_ROW,
            earlier_c,
            f'=IF({q(SHEET_SKILLS)}!B{skills_row}="","",{q(SHEET_SKILLS)}!C{skills_row})',
        )
        cat.font = font(9, bold=True, color=WHITE)
        cat.fill = fill(TEAL)
        cat.alignment = align("center")
        cat.border = THIN
        for col in range(earlier_c, change_c + 1):
            ws.cell(GLANCE_CAT_ROW, col).fill = fill(TEAL)
            ws.cell(GLANCE_CAT_ROW, col).border = THIN

        ws.merge_cells(
            start_row=GLANCE_SKILL_ROW, start_column=earlier_c, end_row=GLANCE_SKILL_ROW, end_column=change_c
        )
        name = ws.cell(
            GLANCE_SKILL_ROW,
            earlier_c,
            f'=IF({q(SHEET_SKILLS)}!B{skills_row}="","",{q(SHEET_SKILLS)}!B{skills_row})',
        )
        name.font = font(11, bold=True, color=NAVY)
        name.fill = fill(PAPER)
        name.alignment = align("center", wrap=True)
        name.border = THIN
        for col in range(earlier_c, change_c + 1):
            ws.cell(GLANCE_SKILL_ROW, col).fill = fill(PAPER)
            ws.cell(GLANCE_SKILL_ROW, col).border = THIN

        skill_name_cell = f"{get_column_letter(earlier_c)}${GLANCE_SKILL_ROW}"
        ly_h = ws.cell(GLANCE_SUB_ROW, earlier_c, "=ViewEarlierYear")
        ty_h = ws.cell(GLANCE_SUB_ROW, later_c, "=ViewLaterYear")
        ch_h = ws.cell(GLANCE_SUB_ROW, change_c, "Change")
        ly_h.fill = fill("5B8FB9")
        ty_h.fill = fill(TEAL)
        ch_h.fill = fill("8A8178")
        for cell in (ly_h, ty_h, ch_h):
            cell.font = font(10, bold=True, color=WHITE)
            cell.alignment = align("center")
            cell.border = THIN
            cell.number_format = "0"

        ws.column_dimensions[get_column_letter(earlier_c)].width = 11
        ws.column_dimensions[get_column_letter(later_c)].width = 11
        ws.column_dimensions[get_column_letter(change_c)].width = 10

        for i in range(NUM_EMPLOYEE_SLOTS):
            r = GLANCE_START + i
            emp_row = 6 + i
            if n == 1:
                ws.cell(r, 1, f'=IF({q(SHEET_EMP)}!B{emp_row}="","",{q(SHEET_EMP)}!B{emp_row})')
                ws.cell(r, 2, f'=IF(A{r}="","",{q(SHEET_EMP)}!C{emp_row})')
                ws.cell(r, 1).font = font(11, bold=True)
                ws.cell(r, 2).font = font(10, color=MUTED)
                ws.cell(r, 1).border = THIN
                ws.cell(r, 2).border = THIN
                ws.row_dimensions[r].height = 22

            def pull(year_name: str) -> str:
                return (
                    f'=IF(OR($A{r}="",{skill_name_cell}="",{year_name}=""),"",'
                    f"IFERROR(1/(1/SUMIFS("
                    f"INDEX({year_block},0,MATCH({year_name},{year_hdr},0)),"
                    f"{emp_rng},$A{r},{skill_rng},{skill_name_cell}"
                    f'))),""))'
                )

            earlier_cell = ws.cell(r, earlier_c, pull("ViewEarlierYear"))
            later_cell = ws.cell(r, later_c, pull("ViewLaterYear"))
            change_cell = ws.cell(
                r,
                change_c,
                f'=IF(OR({get_column_letter(earlier_c)}{r}="",{get_column_letter(later_c)}{r}=""),"",'
                f"{get_column_letter(later_c)}{r}-{get_column_letter(earlier_c)}{r})",
            )
            earlier_cell.fill = fill(LY_FILL)
            later_cell.fill = fill(TY_FILL)
            change_cell.fill = fill(CHANGE_FILL)
            for cell in (earlier_cell, later_cell, change_cell):
                cell.alignment = align("center")
                cell.border = THIN
                cell.font = font(12, bold=True)
            change_cell.font = font(11)
            change_cell.number_format = "+0;-0;0"

    for n in range(1, NUM_SKILL_SLOTS + 1):
        earlier_c, later_c, change_c = glance_skill_columns(n)
        apply_rating_cf(
            ws,
            f"{get_column_letter(earlier_c)}{GLANCE_START}:{get_column_letter(later_c)}{GLANCE_END}",
        )
        apply_change_cf(
            ws,
            f"{get_column_letter(change_c)}{GLANCE_START}:{get_column_letter(change_c)}{GLANCE_END}",
        )

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 16
    ws.row_dimensions[GLANCE_CAT_ROW].height = 18
    ws.row_dimensions[GLANCE_SKILL_ROW].height = 32
    ws.row_dimensions[GLANCE_SUB_ROW].height = 20
    ws.row_dimensions[GLANCE_PAIR_ROW].height = 28
    ws.freeze_panes = f"C{GLANCE_START}"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = NAVY
    apply_print(ws, landscape=True)
    ws.print_title_rows = "1:6"
    ws.print_title_cols = "A:B"


# ===========================================================================
# Ratings (all years — source of truth for data entry)
# ===========================================================================
def build_ratings(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_RATINGS)
    last_col = SUMMARY_FIRST_COL + 4
    add_navigation(ws, SHEET_RATINGS, last_col)

    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=8)
    style_title(ws["A2"], "Ratings — all years")
    ws.merge_cells(start_row=TITLE_ROW, start_column=9, end_row=TITLE_ROW, end_column=last_col)
    ws.cell(
        TITLE_ROW,
        9,
        "Type 1–5 here for any year. Skill Matrix is the two-year glance of the whole crew. "
        "Pick a year pair there to scroll everyone at once — no macros.",
    ).font = font(10, italic=True, color=MUTED)
    ws.row_dimensions[TITLE_ROW].height = 28

    ws.merge_cells(start_row=MATRIX_INFO_ROW, start_column=1, end_row=MATRIX_INFO_ROW, end_column=last_col)
    ws.cell(
        MATRIX_INFO_ROW,
        1,
        '="Rating 1–5. Current calendar year: "&CalendarYear&'
        '". Latest year with ratings: "&IF(LatestYear="","—",LatestYear)&'
        '". Filter the Employee column to focus on one person."',
    ).font = font(11, color=MUTED)

    headers = ["Employee", "Department", "Skill", "Category"]
    for col, label in enumerate(headers, 1):
        cell = ws.cell(MATRIX_HEADER_ROW, col, label)
        cell.font = font(11, bold=True, color=WHITE)
        cell.fill = fill(NAVY)
        cell.alignment = align("center", wrap=True)
        cell.border = THIN

    for slot in range(NUM_YEAR_SLOTS):
        col = year_col(slot)
        cell = ws.cell(MATRIX_HEADER_ROW, col, f"=IF({q(SHEET_YEARS)}!B{8 + slot}=\"\",\"\",{q(SHEET_YEARS)}!B{8 + slot})")
        cell.font = font(11, bold=True, color=WHITE)
        cell.fill = fill("5B8FB9")
        cell.alignment = align("center")
        cell.border = THIN
        cell.number_format = "0"
        ws.column_dimensions[get_column_letter(col)].width = 9

    summary_headers = ["First", "Latest", "vs last year", "Since joined", "Years rated"]
    for i, label in enumerate(summary_headers):
        cell = ws.cell(MATRIX_HEADER_ROW, SUMMARY_FIRST_COL + i, label)
        cell.font = font(10, bold=True, color=WHITE)
        cell.fill = fill("8A8178")
        cell.alignment = align("center", wrap=True)
        cell.border = THIN

    year_start = get_column_letter(YEAR_FIRST_COL)
    year_end = get_column_letter(YEAR_FIRST_COL + NUM_YEAR_SLOTS - 1)
    header_years = f"${year_start}${MATRIX_HEADER_ROW}:${year_end}${MATRIX_HEADER_ROW}"

    # Highlight current calendar year column header
    header_year_range = f"{year_start}{MATRIX_HEADER_ROW}:{year_end}{MATRIX_HEADER_ROW}"
    ws.conditional_formatting.add(
        header_year_range,
        FormulaRule(
            formula=[f"{year_start}{MATRIX_HEADER_ROW}=CalendarYear"],
            fill=fill(TEAL),
            font=font(11, bold=True, color=WHITE),
        ),
    )

    for emp_i in range(NUM_EMPLOYEE_SLOTS):
        emp_row = 6 + emp_i
        for skill_i in range(NUM_SKILL_SLOTS):
            r = matrix_row(emp_i, skill_i)
            skill_row = 6 + skill_i
            occupied = (
                f'AND({q(SHEET_EMP)}!B{emp_row}<>"",{q(SHEET_SKILLS)}!B{skill_row}<>"")'
            )
            ws.cell(r, 1, f'=IF({occupied},{q(SHEET_EMP)}!B{emp_row},"")')
            ws.cell(r, 2, f'=IF(A{r}="","",{q(SHEET_EMP)}!C{emp_row})')
            ws.cell(r, 3, f'=IF(A{r}="","",{q(SHEET_SKILLS)}!B{skill_row})')
            ws.cell(r, 4, f'=IF(A{r}="","",{q(SHEET_SKILLS)}!C{skill_row})')
            ws.cell(r, 1).font = font(11, bold=True)
            for c in range(1, 5):
                ws.cell(r, c).border = THIN
                ws.cell(r, c).font = font(11, bold=c == 1)
            ws.cell(r, 2).font = font(10, color=MUTED)
            ws.cell(r, 4).font = font(10, color=MUTED)

            for slot in range(NUM_YEAR_SLOTS):
                col = year_col(slot)
                cell = ws.cell(r, col, "")
                cell.alignment = align("center")
                cell.border = THIN
                cell.font = font(12, bold=True)
                year = PRESET_YEARS[slot] if slot < len(PRESET_YEARS) else None
                value = rating_for(emp_i, skill_i, year) if year else None
                if value is not None:
                    cell.value = value
                cell.fill = fill(TY_FILL if year == 2026 else LY_FILL)

            rng = f"{year_start}{r}:{year_end}{r}"
            first_c = ws.cell(
                r,
                SUMMARY_FIRST_COL,
                f'=IF(A{r}="","",IFERROR(INDEX({rng},MATCH(TRUE,INDEX({rng}<>"",0),0)),""))',
            )
            latest_c = ws.cell(
                r,
                SUMMARY_FIRST_COL + 1,
                f'=IF(A{r}="","",IFERROR(LOOKUP(2,1/({rng}<>""),{rng}),""))',
            )
            vs_last = ws.cell(
                r,
                SUMMARY_FIRST_COL + 2,
                f'=IF(OR(A{r}="",PreviousYear=""),"",IFERROR('
                f"LOOKUP(2,1/({rng}<>\"\"),{rng})-INDEX({rng},MATCH(PreviousYear,{header_years},0)),\"\"))",
            )
            since = ws.cell(
                r,
                SUMMARY_FIRST_COL + 3,
                f'=IF(OR(A{r}="",{get_column_letter(SUMMARY_FIRST_COL)}{r}="",'
                f'{get_column_letter(SUMMARY_FIRST_COL + 1)}{r}=""),"",'
                f"{get_column_letter(SUMMARY_FIRST_COL + 1)}{r}-{get_column_letter(SUMMARY_FIRST_COL)}{r})",
            )
            counted = ws.cell(
                r,
                SUMMARY_FIRST_COL + 4,
                f'=IF(A{r}="","",COUNT({rng}))',
            )
            for cell in (first_c, latest_c, vs_last, since, counted):
                cell.alignment = align("center")
                cell.border = THIN
                cell.fill = fill(CHANGE_FILL)
            vs_last.number_format = "+0;-0;0"
            since.number_format = "+0;-0;0"

    year_data = f"{year_start}{MATRIX_DATA_START}:{year_end}{MATRIX_DATA_END}"
    add_rating_validation(ws, year_data)
    apply_rating_cf(ws, year_data)
    apply_change_cf(
        ws,
        f"{get_column_letter(SUMMARY_FIRST_COL + 2)}{MATRIX_DATA_START}:{get_column_letter(SUMMARY_FIRST_COL + 3)}{MATRIX_DATA_END}",
    )

    # Highlight current year input cells
    ws.conditional_formatting.add(
        year_data,
        FormulaRule(
            formula=[f'AND({year_start}${MATRIX_HEADER_ROW}=CalendarYear,{year_start}{MATRIX_DATA_START}="")'],
            fill=fill(YELLOW),
        ),
    )

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 14
    for i in range(5):
        ws.column_dimensions[get_column_letter(SUMMARY_FIRST_COL + i)].width = 13
    ws.row_dimensions[MATRIX_HEADER_ROW].height = 28
    ws.freeze_panes = f"E{MATRIX_DATA_START}"
    ws.auto_filter.ref = f"A{MATRIX_HEADER_ROW}:{get_column_letter(last_col)}{MATRIX_DATA_END}"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = "5B8FB9"
    apply_print(ws, landscape=True)
    ws.print_title_rows = "1:5"
    ws.print_title_cols = "A:D"


# ===========================================================================
# Hidden Data
# ===========================================================================
def build_data(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_DATA)
    headers = ["Employee", "Department", "Job Title", "Skill", "Category", "Year", "Rating", "HireDate"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = font(11, bold=True, color=WHITE)
        c.fill = fill(NAVY)

    row = 2
    for emp_i in range(NUM_EMPLOYEE_SLOTS):
        emp_row = 6 + emp_i
        for skill_i in range(NUM_SKILL_SLOTS):
            mrow = matrix_row(emp_i, skill_i)
            for slot in range(NUM_YEAR_SLOTS):
                ycol = get_column_letter(year_col(slot))
                occupied = (
                    f'AND({q(SHEET_RATINGS)}!A{mrow}<>"",{q(SHEET_RATINGS)}!{ycol}{MATRIX_HEADER_ROW}<>"")'
                )
                ws.cell(row, 1, f"=IF({occupied},{q(SHEET_RATINGS)}!A{mrow},\"\")")
                ws.cell(row, 2, f'=IF(A{row}="","",{q(SHEET_EMP)}!C{emp_row})')
                ws.cell(row, 3, f'=IF(A{row}="","",{q(SHEET_EMP)}!D{emp_row})')
                ws.cell(row, 4, f'=IF(A{row}="","",{q(SHEET_RATINGS)}!C{mrow})')
                ws.cell(row, 5, f'=IF(A{row}="","",{q(SHEET_RATINGS)}!D{mrow})')
                ws.cell(row, 6, f"=IF({occupied},{q(SHEET_RATINGS)}!{ycol}{MATRIX_HEADER_ROW},\"\")")
                ws.cell(
                    row,
                    7,
                    f"=IF(OR(A{row}=\"\",{q(SHEET_RATINGS)}!{ycol}{mrow}=\"\"),\"\",{q(SHEET_RATINGS)}!{ycol}{mrow})",
                )
                ws.cell(row, 8, f'=IF(A{row}="","",{q(SHEET_EMP)}!F{emp_row})')
                row += 1

    last = row - 1
    add_defined_name(wb, "DataEmployee", f"{q(SHEET_DATA)}!$A$2:$A${last}")
    add_defined_name(wb, "DataDept", f"{q(SHEET_DATA)}!$B$2:$B${last}")
    add_defined_name(wb, "DataSkill", f"{q(SHEET_DATA)}!$D$2:$D${last}")
    add_defined_name(wb, "DataYear", f"{q(SHEET_DATA)}!$F$2:$F${last}")
    add_defined_name(wb, "DataRating", f"{q(SHEET_DATA)}!$G$2:$G${last}")
    ws.sheet_state = "hidden"
    for i, w in enumerate([20, 14, 20, 22, 14, 10, 10, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ===========================================================================
# Calc (hidden helpers for dashboard + charts)
# ===========================================================================
def build_calc(wb: Workbook) -> None:
    ws = wb[SHEET_CALC] if SHEET_CALC in wb.sheetnames else wb.create_sheet(SHEET_CALC)

    ws["A1"] = "LatestYear"
    ws["B1"] = '=IFERROR(MAXIFS(DataYear,DataRating,">=1"),YEAR(TODAY()))'
    ws["A2"] = "PreviousYear"
    ws["B2"] = '=IFERROR(MAXIFS(DataYear,DataRating,">=1",DataYear,"<"&B1),"")'
    add_defined_name(wb, "LatestYear", f"{q(SHEET_CALC)}!$B$1")
    add_defined_name(wb, "PreviousYear", f"{q(SHEET_CALC)}!$B$2")
    add_defined_name(wb, "EmpFilter", f"{q(SHEET_DASH)}!$C$5")
    add_defined_name(wb, "SkillSetFilter", f"{q(SHEET_DASH)}!$F$5")
    add_defined_name(wb, "DeptFilter", f"{q(SHEET_DASH)}!$I$5")

    # Consecutive year pairs for the Skill Matrix scroller (no VBA)
    ws["G30"] = "YearPair"
    for i in range(NUM_YEAR_SLOTS - 1):
        r = 31 + i
        ws.cell(
            r,
            7,
            f'=IF(OR({q(SHEET_YEARS)}!B{8 + i}="",{q(SHEET_YEARS)}!B{9 + i}=""),"",'
            f'{q(SHEET_YEARS)}!B{8 + i}&" vs "&{q(SHEET_YEARS)}!B{9 + i})',
        )
    pair_last = 31 + NUM_YEAR_SLOTS - 2
    add_defined_name(
        wb,
        "YearPairList",
        f"{q(SHEET_CALC)}!$G$31:INDEX({q(SHEET_CALC)}!$G$31:$G${pair_last},COUNTA({q(SHEET_CALC)}!$G$31:$G${pair_last}))",
    )
    pair_dv = DataValidation(type="list", formula1="=YearPairList", allow_blank=False)
    pair_dv.promptTitle = "Year pair"
    pair_dv.prompt = "Pick two years. Every employee on Skill Matrix updates together."
    wb[SHEET_MATRIX].add_data_validation(pair_dv)
    pair_dv.add("B3")

    # Team average by year (respects department filter)
    ws["A4"] = "Year"
    ws["B4"] = "TeamAvg"
    for slot in range(NUM_YEAR_SLOTS):
        r = 5 + slot
        ws.cell(r, 1, f"=IF({q(SHEET_YEARS)}!B{8 + slot}=\"\",\"\",{q(SHEET_YEARS)}!B{8 + slot})")
        ws.cell(
            r,
            2,
            f'=IF(A{r}="","",IFERROR(IF(DeptFilter="(All departments)",'
            f"AVERAGEIFS(DataRating,DataYear,A{r}),"
            f"AVERAGEIFS(DataRating,DataYear,A{r},DataDept,DeptFilter)),NA()))",
        )
        ws.cell(r, 1).number_format = "0"
        ws.cell(r, 2).number_format = "0.00"

    # Selected employee average by year
    ws["D4"] = "Year"
    ws["E4"] = "EmpAvg"
    for slot in range(NUM_YEAR_SLOTS):
        r = 5 + slot
        ws.cell(r, 4, f"=A{r}")
        ws.cell(
            r,
            5,
            f'=IF(OR(D{r}="",EmpFilter=""),"",IFERROR(AVERAGEIFS(DataRating,DataYear,D{r},DataEmployee,EmpFilter),NA()))',
        )
        ws.cell(r, 5).number_format = "0.00"

    # Employee skill compare: First / Previous / Latest
    ws["G4"] = "Skill"
    ws["H4"] = "First"
    ws["I4"] = "LastYear"
    ws["J4"] = "Latest"
    for skill_i in range(NUM_SKILL_SLOTS):
        r = 5 + skill_i
        srow = 6 + skill_i
        ws.cell(r, 7, f'=IF({q(SHEET_SKILLS)}!B{srow}="","",{q(SHEET_SKILLS)}!B{srow})')
        ws.cell(
            r,
            8,
            f'=IF(OR(G{r}="",EmpFilter=""),"",IF(COUNTIFS(DataEmployee,EmpFilter,DataSkill,G{r},DataRating,">=1")=0,"",'
            f'IFERROR(SUMIFS(DataRating,DataEmployee,EmpFilter,DataSkill,G{r},DataYear,'
            f'MINIFS(DataYear,DataEmployee,EmpFilter,DataSkill,G{r},DataRating,">=1")),"")))',
        )
        ws.cell(
            r,
            9,
            f'=IF(OR(G{r}="",EmpFilter="",PreviousYear=""),"",IF(COUNTIFS(DataEmployee,EmpFilter,DataSkill,G{r},DataYear,PreviousYear,DataRating,">=1")=0,"",'
            f'IFERROR(SUMIFS(DataRating,DataEmployee,EmpFilter,DataSkill,G{r},DataYear,PreviousYear),"")))',
        )
        ws.cell(
            r,
            10,
            f'=IF(OR(G{r}="",EmpFilter=""),"",IF(COUNTIFS(DataEmployee,EmpFilter,DataSkill,G{r},DataYear,LatestYear,DataRating,">=1")=0,"",'
            f'IFERROR(SUMIFS(DataRating,DataEmployee,EmpFilter,DataSkill,G{r},DataYear,LatestYear),"")))',
        )

    # Skill-set required skills
    ws["L4"] = "SetSkill"
    ws["L5"] = "Min"
    for k in range(MAX_SET_SKILLS):
        col = 13 + k  # M onwards
        ws.cell(4, col, f'=IFERROR(INDEX(FILTER({q(SHEET_SETS)}!$B$6:$B${5 + NUM_SET_SLOTS},{q(SHEET_SETS)}!$A$6:$A${5 + NUM_SET_SLOTS}=SkillSetFilter),{k + 1}),"")')
        ws.cell(5, col, f'=IFERROR(INDEX(FILTER({q(SHEET_SETS)}!$C$6:$C${5 + NUM_SET_SLOTS},{q(SHEET_SETS)}!$A$6:$A${5 + NUM_SET_SLOTS}=SkillSetFilter),{k + 1}),"")')

    # Per-employee skill-set scores
    ws["L7"] = "Employee"
    ws["M7"] = "Score"
    ws["N7"] = "VsLast"
    ws["O7"] = "Coverage"
    ws["P7"] = "MetMin"
    ws["Q7"] = "Rank"
    skill_cols = [get_column_letter(13 + k) for k in range(MAX_SET_SKILLS)]
    for emp_i in range(NUM_EMPLOYEE_SLOTS):
        r = 8 + emp_i
        emp_row = 6 + emp_i
        ws.cell(r, 12, f'=IF({q(SHEET_EMP)}!B{emp_row}="","",{q(SHEET_EMP)}!B{emp_row})')
        latest_parts = []
        prev_parts = []
        cov_parts = []
        min_parts = []
        for k, col in enumerate(skill_cols):
            latest_parts.append(
                f'IF({col}$4="","",SUMIFS(DataRating,DataEmployee,L{r},DataSkill,{col}$4,DataYear,LatestYear))'
            )
            prev_parts.append(
                f'IF(OR({col}$4="",PreviousYear=""),"",SUMIFS(DataRating,DataEmployee,L{r},DataSkill,{col}$4,DataYear,PreviousYear))'
            )
            cov_parts.append(
                f'IF(OR({col}$4="",SUMIFS(DataRating,DataEmployee,L{r},DataSkill,{col}$4,DataYear,LatestYear)=0),0,1)'
            )
            min_parts.append(
                f'IF(OR({col}$4="",{col}$5=""),0,IF(SUMIFS(DataRating,DataEmployee,L{r},DataSkill,{col}$4,DataYear,LatestYear)>={col}$5,1,0))'
            )
        ws.cell(
            r,
            13,
            f'=IF(L{r}="","",IFERROR(IF(({"+".join(cov_parts)})=0,"",'
            f'({"+".join(latest_parts)})/({"+".join(cov_parts)})),""))',
        )
        ws.cell(
            r,
            14,
            f'=IF(OR(L{r}="",M{r}="",PreviousYear=""),"",IFERROR('
            f'IF(({"+".join(cov_parts)})=0,"",({"+".join(latest_parts)})/({"+".join(cov_parts)}))'
            f'-IFERROR(({"+".join(prev_parts)})/MAX(1,{"+".join(cov_parts)}),""),""))',
        )
        ws.cell(r, 15, f'=IF(L{r}="","",{"+".join(cov_parts)}&"/"&COUNTA({skill_cols[0]}$4:{skill_cols[-1]}$4))')
        ws.cell(r, 16, f'=IF(L{r}="","",{"+".join(min_parts)}&"/"&COUNT({skill_cols[0]}$5:{skill_cols[-1]}$5))')
        ws.cell(
            r,
            17,
            f'=IF(OR(L{r}="",M{r}=""),"",RANK(M{r},$M$8:$M${7 + NUM_EMPLOYEE_SLOTS},0)+COUNTIF($M$8:M{r},M{r})-1)',
        )
        ws.cell(r, 13).number_format = "0.00"
        ws.cell(r, 14).number_format = "+0.00;-0.00;0.00"

    # Ranked output for chart (fixed 16 rows)
    ws["S7"] = "RankedEmployee"
    ws["T7"] = "RankedScore"
    for i in range(NUM_EMPLOYEE_SLOTS):
        r = 8 + i
        ws.cell(r, 19, f'=IFERROR(INDEX($L$8:$L${7 + NUM_EMPLOYEE_SLOTS},MATCH({i + 1},$Q$8:$Q${7 + NUM_EMPLOYEE_SLOTS},0)),"")')
        ws.cell(r, 20, f'=IFERROR(INDEX($M$8:$M${7 + NUM_EMPLOYEE_SLOTS},MATCH({i + 1},$Q$8:$Q${7 + NUM_EMPLOYEE_SLOTS},0)),"")')
        ws.cell(r, 20).number_format = "0.00"

    ws.sheet_state = "hidden"


# ===========================================================================
# Dashboard
# ===========================================================================
def kpi_card(ws, col, title, formula, color, fmt="0.00") -> None:
    ws.merge_cells(start_row=7, start_column=col, end_row=7, end_column=col + 2)
    ws.merge_cells(start_row=8, start_column=col, end_row=9, end_column=col + 2)
    head = ws.cell(7, col, title)
    head.font = font(10, bold=True, color=WHITE)
    head.fill = fill(color)
    head.alignment = align("center")
    val = ws.cell(8, col, formula)
    val.font = font(22, bold=True, color=NAVY if color == GOLD else WHITE)
    val.fill = fill(YELLOW if color == GOLD else color)
    val.alignment = align("center")
    val.number_format = fmt
    for r in (7, 8, 9):
        for c in range(col, col + 3):
            ws.cell(r, c).fill = fill(YELLOW if color == GOLD and r > 7 else color)
            ws.cell(r, c).border = Border(
                left=Side(style="thin", color=WHITE),
                right=Side(style="thin", color=WHITE),
            )


def build_dashboard(wb: Workbook) -> None:
    ws = wb.create_sheet(SHEET_DASH, 1)
    add_navigation(ws, SHEET_DASH, 18)

    ws.merge_cells("A2:R2")
    style_title(ws["A2"], "Dashboard — year on year")
    ws.merge_cells("A3:R3")
    ws["A3"] = (
        '="Team view, one person, or a job profile. Latest year with ratings: "&IF(LatestYear="","—",LatestYear)'
        '&"  ·  Previous: "&IF(PreviousYear="","—",PreviousYear)&"  ·  Calendar: "&CalendarYear'
    )
    ws["A3"].font = font(11, italic=True, color=MUTED)

    # ----- Filter bar -----
    ws.merge_cells("A4:R4")
    ws["A4"] = "Filters — pick an employee to explore their history, a skill set to see who can do that job, and a department to focus the team charts."
    ws["A4"].font = font(10, italic=True, color=WHITE)
    ws["A4"].fill = fill(NAVY)
    for c in range(1, 19):
        ws.cell(4, c).fill = fill(NAVY)

    ws["A5"] = "Employee"
    ws["B5"] = "→"
    ws["C5"] = EMPLOYEES[0][1]
    ws["D5"] = "Skill set"
    ws["E5"] = "→"
    ws["F5"] = SKILL_SETS[0][0]
    ws["G5"] = "Department"
    ws["H5"] = "→"
    ws["I5"] = "(All departments)"
    ws["K5"] = '=HYPERLINK("#\'Years\'!A1","Add / manage years →")'
    ws["K5"].font = font(11, bold=True, color=WHITE, underline="single")
    ws["K5"].fill = fill(TEAL)
    ws.merge_cells("K5:M5")
    ws["K5"].alignment = align("center")
    for col in (1, 4, 7):
        ws.cell(5, col).font = font(11, bold=True, color=WHITE)
        ws.cell(5, col).fill = fill(TEAL)
        ws.cell(5, col).alignment = align("center")
    for col in (2, 5, 8):
        ws.cell(5, col).fill = fill(TEAL)
        ws.cell(5, col).font = font(11, bold=True, color=WHITE)
        ws.cell(5, col).alignment = align("center")
    for col in (3, 6, 9):
        style_input(ws.cell(5, col))
        ws.cell(5, col).font = font(12, bold=True, color=NAVY)

    emp_dv = DataValidation(type="list", formula1="=EmployeeList", allow_blank=False)
    emp_dv.promptTitle = "Employee"
    emp_dv.prompt = "Choose someone to see their skill history"
    ws.add_data_validation(emp_dv)
    emp_dv.add("C5")

    set_dv = DataValidation(type="list", formula1="=SkillSetList", allow_blank=False)
    set_dv.promptTitle = "Skill set"
    set_dv.prompt = "Choose a job profile to rank who can do it"
    ws.add_data_validation(set_dv)
    set_dv.add("F5")

    dept_dv = DataValidation(type="list", formula1="=DeptFilterList", allow_blank=False)
    dept_dv.promptTitle = "Department"
    dept_dv.prompt = "Limit the team overview to one department, or All"
    ws.add_data_validation(dept_dv)
    dept_dv.add("I5")

    ws.row_dimensions[5].height = 26

    # ----- Org overview KPIs -----
    ws.merge_cells("A6:R6")
    ws["A6"] = '="Team overview"&IF(DeptFilter="(All departments)",""," — "&DeptFilter)'
    ws["A6"].font = font(14, bold=True, color=NAVY)

    kpi_card(
        ws,
        1,
        "Latest year average",
        '=IFERROR(ROUND(IF(DeptFilter="(All departments)",AVERAGEIFS(DataRating,DataYear,LatestYear),AVERAGEIFS(DataRating,DataYear,LatestYear,DataDept,DeptFilter)),2),"—")',
        TEAL,
    )
    kpi_card(ws, 4, "Previous year average",
             '=IFERROR(ROUND(IF(DeptFilter="(All departments)",AVERAGEIFS(DataRating,DataYear,PreviousYear),AVERAGEIFS(DataRating,DataYear,PreviousYear,DataDept,DeptFilter)),2),"—")',
             "5B8FB9")
    kpi_card(ws, 7, "Change vs last year",
             '=IFERROR(ROUND(IF(DeptFilter="(All departments)",AVERAGEIFS(DataRating,DataYear,LatestYear),AVERAGEIFS(DataRating,DataYear,LatestYear,DataDept,DeptFilter))'
             '-IF(DeptFilter="(All departments)",AVERAGEIFS(DataRating,DataYear,PreviousYear),AVERAGEIFS(DataRating,DataYear,PreviousYear,DataDept,DeptFilter)),2),"—")',
             GOLD, "+0.00;-0.00;0.00")
    kpi_card(
        ws,
        10,
        "Change vs first year",
        '=IFERROR(ROUND(IF(DeptFilter="(All departments)",AVERAGEIFS(DataRating,DataYear,LatestYear),AVERAGEIFS(DataRating,DataYear,LatestYear,DataDept,DeptFilter))'
        '-IF(DeptFilter="(All departments)",AVERAGEIFS(DataRating,DataYear,MINIFS(DataYear,DataRating,">=1")),AVERAGEIFS(DataRating,DataYear,MINIFS(DataYear,DataRating,">=1",DataDept,DeptFilter),DataDept,DeptFilter)),2),"—")',
        NAVY,
        "+0.00;-0.00;0.00",
    )
    kpi_card(ws, 13, "Ratings in latest year",
             '=IF(DeptFilter="(All departments)",COUNTIFS(DataYear,LatestYear,DataRating,">=1"),COUNTIFS(DataYear,LatestYear,DataRating,">=1",DataDept,DeptFilter))',
             "6B5B95", "0")

    # Line chart: team by year
    chart1 = LineChart()
    chart1.title = "Team average — year on year"
    chart1.y_axis.title = "Average (1–5)"
    chart1.y_axis.scaling.min = 1
    chart1.y_axis.scaling.max = 5
    chart1.style = 10
    chart1.height = 8
    chart1.width = 15
    chart1.legend = None
    data = Reference(wb[SHEET_CALC], min_col=2, min_row=4, max_row=4 + NUM_YEAR_SLOTS)
    cats = Reference(wb[SHEET_CALC], min_col=1, min_row=5, max_row=4 + NUM_YEAR_SLOTS)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    if chart1.series:
        chart1.series[0].graphicalProperties.line.solidFill = TEAL
        chart1.series[0].graphicalProperties.line.width = 25000
        chart1.series[0].marker = Marker(symbol="circle", size=8)
    ws.add_chart(chart1, "A11")

    # Skill latest vs last vs first table
    ws["I11"] = "Skill"
    ws["J11"] = "First year"
    ws["K11"] = "Last year"
    ws["L11"] = "Latest"
    ws["M11"] = "vs last year"
    for col in range(9, 14):
        c = ws.cell(11, col)
        c.font = font(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center", wrap=True)
        c.border = THIN
    for skill_i in range(len(SKILLS)):
        r = 12 + skill_i
        srow = 6 + skill_i
        ws.cell(r, 9, f'=IF({q(SHEET_SKILLS)}!B{srow}="","",{q(SHEET_SKILLS)}!B{srow})')
        ws.cell(
            r,
            10,
            f'=IF(I{r}="","",IFERROR(IF(DeptFilter="(All departments)",'
            f'AVERAGEIFS(DataRating,DataSkill,I{r},DataYear,MINIFS(DataYear,DataSkill,I{r},DataRating,">=1")),'
            f'AVERAGEIFS(DataRating,DataSkill,I{r},DataYear,MINIFS(DataYear,DataSkill,I{r},DataRating,">=1"),DataDept,DeptFilter)),""))',
        )
        ws.cell(
            r,
            11,
            f'=IF(OR(I{r}="",PreviousYear=""),"",IFERROR(IF(DeptFilter="(All departments)",'
            f"AVERAGEIFS(DataRating,DataSkill,I{r},DataYear,PreviousYear),"
            f"AVERAGEIFS(DataRating,DataSkill,I{r},DataYear,PreviousYear,DataDept,DeptFilter)),\"\"))",
        )
        ws.cell(
            r,
            12,
            f'=IF(I{r}="","",IFERROR(IF(DeptFilter="(All departments)",'
            f"AVERAGEIFS(DataRating,DataSkill,I{r},DataYear,LatestYear),"
            f"AVERAGEIFS(DataRating,DataSkill,I{r},DataYear,LatestYear,DataDept,DeptFilter)),\"\"))",
        )
        ws.cell(r, 13, f'=IF(OR(K{r}="",L{r}=""),"",ROUND(L{r}-K{r},2))')
        for c in range(9, 14):
            ws.cell(r, c).border = THIN
            ws.cell(r, c).font = font(10)
            if c > 9:
                ws.cell(r, c).alignment = align("center")
                ws.cell(r, c).number_format = "0.00" if c < 13 else "+0.00;-0.00;0.00"
        apply_change_cf(ws, f"M{r}")

    chart_skill = BarChart()
    chart_skill.type = "bar"
    chart_skill.grouping = "clustered"
    chart_skill.title = "Each skill: last year vs latest"
    chart_skill.style = 10
    chart_skill.y_axis.title = None
    chart_skill.x_axis.scaling.min = 1
    chart_skill.x_axis.scaling.max = 5
    chart_skill.height = 8
    chart_skill.width = 14
    sdata = Reference(ws, min_col=11, min_row=11, max_col=12, max_row=11 + len(SKILLS))
    scats = Reference(ws, min_col=9, min_row=12, max_row=11 + len(SKILLS))
    chart_skill.add_data(sdata, titles_from_data=True)
    chart_skill.set_categories(scats)
    chart_skill.legend.position = "b"
    if chart_skill.series:
        chart_skill.series[0].graphicalProperties.solidFill = "5B8FB9"
        if len(chart_skill.series) > 1:
            chart_skill.series[1].graphicalProperties.solidFill = TEAL
    ws.add_chart(chart_skill, "O11")

    # ----- Employee explorer -----
    emp_title_row = 24
    ws.merge_cells(start_row=emp_title_row, start_column=1, end_row=emp_title_row, end_column=18)
    ws.cell(
        emp_title_row,
        1,
        '=IF(EmpFilter="","Select an employee in the filter above",'
        'EmpFilter&" — skill overview since they joined")',
    ).font = font(14, bold=True, color=NAVY)

    ws.merge_cells(start_row=25, start_column=1, end_row=25, end_column=18)
    ws["A25"] = (
        '=IF(EmpFilter="","",'
        '"Department: "&IFERROR(INDEX(Employees!C6:C21,MATCH(EmpFilter,Employees!B6:B21,0)),"—")'
        '&"   ·   Title: "&IFERROR(INDEX(Employees!D6:D21,MATCH(EmpFilter,Employees!B6:B21,0)),"—")'
        '&"   ·   Hired: "&TEXT(IFERROR(INDEX(Employees!F6:F21,MATCH(EmpFilter,Employees!B6:B21,0)),""),"YYYY-MM-DD")'
        '&"   ·   First ratings: "&IFERROR(MINIFS(DataYear,DataEmployee,EmpFilter,DataRating,">=1"),"—")'
        '&"   ·   Latest: "&LatestYear)'
    )
    ws["A25"].font = font(11, color=MUTED)

    emp_kpis = [
        (1, "Latest average", '=IFERROR(ROUND(AVERAGEIFS(DataRating,DataEmployee,EmpFilter,DataYear,LatestYear),2),"—")', TEAL, "0.00"),
        (4, "vs last year",
         '=IFERROR(ROUND(AVERAGEIFS(DataRating,DataEmployee,EmpFilter,DataYear,LatestYear)'
         '-AVERAGEIFS(DataRating,DataEmployee,EmpFilter,DataYear,PreviousYear),2),"—")', GOLD, "+0.00;-0.00;0.00"),
        (7, "vs first year (since joined)",
         '=IFERROR(ROUND(AVERAGEIFS(DataRating,DataEmployee,EmpFilter,DataYear,LatestYear)'
         '-AVERAGEIFS(DataRating,DataEmployee,EmpFilter,DataYear,MINIFS(DataYear,DataEmployee,EmpFilter,DataRating,">=1")),2),"—")',
         NAVY, "+0.00;-0.00;0.00"),
        (10, "Skills below target",
         '=COUNTIFS(Calc!G5:G20,"<>",Calc!J5:J20,"<"&TargetRating)', RED, "0"),
        (13, "Skills rated this year",
         '=COUNTIFS(DataEmployee,EmpFilter,DataYear,LatestYear,DataRating,">=1")', "6B5B95", "0"),
    ]
    for col, title, formula, color, fmt in emp_kpis:
        ws.merge_cells(start_row=26, start_column=col, end_row=26, end_column=col + 2)
        ws.merge_cells(start_row=27, start_column=col, end_row=28, end_column=col + 2)
        h = ws.cell(26, col, title)
        h.font = font(9, bold=True, color=WHITE)
        h.fill = fill(color)
        h.alignment = align("center", wrap=True)
        v = ws.cell(27, col, formula)
        v.font = font(20, bold=True, color=NAVY if color in (GOLD, RED) else WHITE)
        v.fill = fill(YELLOW if color in (GOLD, RED) else color)
        v.alignment = align("center")
        v.number_format = fmt
        for r in (26, 27, 28):
            for c in range(col, col + 3):
                ws.cell(r, c).fill = fill(YELLOW if color in (GOLD, RED) and r > 26 else color)
                ws.cell(r, c).border = Border(
                    left=Side(style="thin", color=WHITE),
                    right=Side(style="thin", color=WHITE),
                )

    # Employee skill table
    ws["A30"] = "Skill"
    ws["B30"] = "First (joined)"
    ws["C30"] = "Last year"
    ws["D30"] = "Latest"
    ws["E30"] = "vs last year"
    ws["F30"] = "Since joined"
    for col in range(1, 7):
        c = ws.cell(30, col)
        c.font = font(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center", wrap=True)
        c.border = THIN
    for skill_i in range(NUM_SKILL_SLOTS):
        r = 31 + skill_i
        ws.cell(r, 1, f"=IF(Calc!G{5 + skill_i}=\"\",\"\",Calc!G{5 + skill_i})")
        ws.cell(r, 2, f"=IF(A{r}=\"\",\"\",IF(Calc!H{5 + skill_i}=0,\"\",Calc!H{5 + skill_i}))")
        ws.cell(r, 3, f"=IF(A{r}=\"\",\"\",IF(Calc!I{5 + skill_i}=0,\"\",Calc!I{5 + skill_i}))")
        ws.cell(r, 4, f"=IF(A{r}=\"\",\"\",IF(Calc!J{5 + skill_i}=0,\"\",Calc!J{5 + skill_i}))")
        ws.cell(r, 5, f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})')
        ws.cell(r, 6, f'=IF(OR(B{r}="",D{r}=""),"",D{r}-B{r})')
        for c in range(1, 7):
            ws.cell(r, c).border = THIN
            ws.cell(r, c).alignment = align("center" if c > 1 else "left")
        ws.cell(r, 5).number_format = "+0;-0;0"
        ws.cell(r, 6).number_format = "+0;-0;0"
    apply_rating_cf(ws, f"B31:D{30 + NUM_SKILL_SLOTS}")
    apply_change_cf(ws, f"E31:F{30 + NUM_SKILL_SLOTS}")

    chart_emp_line = LineChart()
    chart_emp_line.title = "This person's average by year"
    chart_emp_line.y_axis.title = "Average (1–5)"
    chart_emp_line.y_axis.scaling.min = 1
    chart_emp_line.y_axis.scaling.max = 5
    chart_emp_line.style = 10
    chart_emp_line.height = 8
    chart_emp_line.width = 12
    chart_emp_line.legend = None
    edata = Reference(wb[SHEET_CALC], min_col=5, min_row=4, max_row=4 + NUM_YEAR_SLOTS)
    ecats = Reference(wb[SHEET_CALC], min_col=4, min_row=5, max_row=4 + NUM_YEAR_SLOTS)
    chart_emp_line.add_data(edata, titles_from_data=True)
    chart_emp_line.set_categories(ecats)
    if chart_emp_line.series:
        chart_emp_line.series[0].graphicalProperties.line.solidFill = NAVY
        chart_emp_line.series[0].marker = Marker(symbol="circle", size=7)
    ws.add_chart(chart_emp_line, "H30")

    chart_emp_bar = BarChart()
    chart_emp_bar.type = "bar"
    chart_emp_bar.grouping = "clustered"
    chart_emp_bar.title = "Each skill: first vs last year vs latest"
    chart_emp_bar.style = 10
    chart_emp_bar.x_axis.scaling.min = 1
    chart_emp_bar.x_axis.scaling.max = 5
    chart_emp_bar.height = 8
    chart_emp_bar.width = 14
    eb = Reference(wb[SHEET_CALC], min_col=8, min_row=4, max_col=10, max_row=4 + len(SKILLS))
    ebcats = Reference(wb[SHEET_CALC], min_col=7, min_row=5, max_row=4 + len(SKILLS))
    chart_emp_bar.add_data(eb, titles_from_data=True)
    chart_emp_bar.set_categories(ebcats)
    chart_emp_bar.legend.position = "b"
    if chart_emp_bar.series:
        chart_emp_bar.series[0].graphicalProperties.solidFill = "8A8178"
        if len(chart_emp_bar.series) > 1:
            chart_emp_bar.series[1].graphicalProperties.solidFill = "5B8FB9"
        if len(chart_emp_bar.series) > 2:
            chart_emp_bar.series[2].graphicalProperties.solidFill = TEAL
    ws.add_chart(chart_emp_bar, "M30")

    # ----- Who can do the job -----
    job_row = 50
    ws.merge_cells(start_row=job_row, start_column=1, end_row=job_row, end_column=18)
    ws.cell(
        job_row,
        1,
        '=IF(SkillSetFilter="","Select a skill set in the filter above",'
        '"Who can do the job — "&SkillSetFilter&"  (ranked by latest-year average on the required skills)")',
    ).font = font(14, bold=True, color=NAVY)

    ws.merge_cells(start_row=51, start_column=1, end_row=51, end_column=12)
    ws["A51"] = (
        f'="Required skills: "&TEXTJOIN(", ",TRUE,FILTER({q(SHEET_SETS)}!B6:B{5 + NUM_SET_SLOTS},{q(SHEET_SETS)}!A6:A{5 + NUM_SET_SLOTS}=SkillSetFilter))'
        f'&"   ·   Minimums: "&TEXTJOIN(", ",TRUE,FILTER({q(SHEET_SETS)}!C6:C{5 + NUM_SET_SLOTS},{q(SHEET_SETS)}!A6:A{5 + NUM_SET_SLOTS}=SkillSetFilter))'
    )
    ws["A51"].font = font(11, color=MUTED)

    headers = ["Rank", "Employee", "Department", "Job title", "Latest avg", "vs last year", "Skills rated", "At / above min"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(52, i, h)
        c.font = font(10, bold=True, color=WHITE)
        c.fill = fill(NAVY)
        c.alignment = align("center", wrap=True)
        c.border = THIN
    ws.row_dimensions[52].height = 28

    for i in range(NUM_EMPLOYEE_SLOTS):
        r = 53 + i
        ws.cell(r, 1, f'=IF(Calc!S{8 + i}="","",{i + 1})')
        ws.cell(r, 2, f"=IF(Calc!S{8 + i}=\"\",\"\",Calc!S{8 + i})")
        ws.cell(
            r,
            3,
            f'=IF(B{r}="","",IFERROR(INDEX(Employees!C6:C21,MATCH(B{r},Employees!B6:B21,0)),""))',
        )
        ws.cell(
            r,
            4,
            f'=IF(B{r}="","",IFERROR(INDEX(Employees!D6:D21,MATCH(B{r},Employees!B6:B21,0)),""))',
        )
        ws.cell(r, 5, f"=IF(B{r}=\"\",\"\",Calc!T{8 + i})")
        ws.cell(
            r,
            6,
            f'=IF(B{r}="","",IFERROR(INDEX(Calc!N8:N23,MATCH(B{r},Calc!L8:L23,0)),""))',
        )
        ws.cell(
            r,
            7,
            f'=IF(B{r}="","",IFERROR(INDEX(Calc!O8:O23,MATCH(B{r},Calc!L8:L23,0)),""))',
        )
        ws.cell(
            r,
            8,
            f'=IF(B{r}="","",IFERROR(INDEX(Calc!P8:P23,MATCH(B{r},Calc!L8:L23,0)),""))',
        )
        for c in range(1, 9):
            ws.cell(r, c).border = THIN
            ws.cell(r, c).alignment = align("center" if c != 2 and c != 4 else "left")
        ws.cell(r, 5).number_format = "0.00"
        ws.cell(r, 6).number_format = "+0.00;-0.00;0.00"
        ws.cell(r, 2).font = font(11, bold=True)
    apply_change_cf(ws, f"F53:F{52 + NUM_EMPLOYEE_SLOTS}")
    ws.conditional_formatting.add(
        "A53:H53",
        FormulaRule(formula=["$B53<>\"\""], fill=fill("D4EDDA"), font=font(11, bold=True)),
    )

    chart_job = BarChart()
    chart_job.type = "bar"
    chart_job.title = "Best current fit for the selected skill set"
    chart_job.style = 10
    chart_job.x_axis.scaling.min = 1
    chart_job.x_axis.scaling.max = 5
    chart_job.height = 8
    chart_job.width = 14
    chart_job.legend = None
    jdata = Reference(wb[SHEET_CALC], min_col=20, min_row=7, max_row=7 + len(EMPLOYEES))
    jcats = Reference(wb[SHEET_CALC], min_col=19, min_row=8, max_row=7 + len(EMPLOYEES))
    chart_job.add_data(jdata, titles_from_data=True)
    chart_job.set_categories(jcats)
    if chart_job.series:
        chart_job.series[0].graphicalProperties.solidFill = TEAL
    ws.add_chart(chart_job, "J52")

    ws.merge_cells("A70:R71")
    ws["A70"] = (
        "Tip: change Employee, Skill set, or Department in the yellow filter cells at the top — "
        "charts and tables update together. Add a year on the Years sheet (or just wait; 2023–2034 are already columns). "
        "Add departments and job titles on Lists, then pick them on Employees."
    )
    ws["A70"].font = font(10, italic=True, color=MUTED)
    ws["A70"].alignment = align("left", wrap=True)

    widths = [16, 16, 16, 16, 14, 14, 14, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col in range(9, 19):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.column_dimensions["A"].width = 22
    ws.freeze_panes = "A6"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = GOLD
    apply_print(ws, landscape=True)


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
    add_navigation(ws, SHEET_HOW, 12)
    ws.merge_cells("A2:L2")
    style_title(ws["A2"], "How to use this skill matrix")
    ws.merge_cells("A3:L3")
    ws["A3"] = (
        "Rate skills from 1 to 5 for a live-event vision crew, keep every year you have rated, "
        "and use the Dashboard to look at the whole team, one operator over time, or who is strongest for a seat. "
        "Sample data is included — replace it with your crew."
    )
    ws["A3"].font = font(12, color=MUTED)
    ws["A3"].alignment = align("left", wrap=True)
    ws.row_dimensions[3].height = 36

    box(
        ws, 5, 1, 14, 6,
        "1. Glance at the whole crew",
        "Open Skill Matrix. Each row is one person. Each skill has two years and a Change column.\n\n"
        "At the top, pick a Year pair (for example 2025 vs 2026). "
        "That dropdown scrolls every employee to those two years at once. No macros.\n\n"
        "Green change = improved. Red = declined.\n\n"
        "Type 1–5 ratings on the Ratings sheet (any year). The glance view updates.",
        NAVY,
    )
    box(
        ws, 5, 7, 14, 12,
        "2. New year — automatic, or add one",
        "Years 2023–2034 are already on Ratings. When a new year starts, open Ratings and fill that column. "
        "Then pick the new pair on Skill Matrix (for example 2026 vs 2027).\n\n"
        "The Years sheet shows whether this calendar year is in the list. "
        "To add 2035 or later: type it in the next yellow row on Years, oldest to newest. "
        "A new column appears on Ratings, and a new pair appears in the dropdown.",
        TEAL,
    )
    box(
        ws, 16, 1, 24, 6,
        "3. Add people, skills, departments, titles",
        "Employees — next yellow row. Pick Department and Job title from the dropdowns. Fill Hire date for “since they joined”.\n\n"
        "Skills — next yellow row. Every person gets a row for the new skill.\n\n"
        "Lists — add a department or job title in a yellow row. It shows up in the dropdowns. Add it here first, then choose it on Employees.",
        "3D5A80",
    )
    box(
        ws, 16, 7, 24, 12,
        "4. Dashboard filters",
        "Employee — that person’s skills, last year vs latest, and since they joined (first recorded rating).\n\n"
        "Skill set — ranks who currently looks strongest for that seat "
        "(Barco Vision Switching, Camera Switching, Content Operation, Systems Tech, "
        "Technical Setup, Live Streaming, or Troubleshooting).\n\n"
        "Department — Vision, Content, Systems, Streaming, or All departments.\n\n"
        "Line charts show year on year on year, not just two years.",
        GOLD,
    )

    ws.merge_cells("A26:L26")
    ws["A26"] = "Rating scale"
    ws["A26"].font = font(14, bold=True, color=NAVY)
    for i, (score, level, meaning) in enumerate(RATING_SCALE):
        c1 = 1 + i * 2
        c2 = c1 + 1
        ws.merge_cells(start_row=27, start_column=c1, end_row=27, end_column=c2)
        ws.merge_cells(start_row=28, start_column=c1, end_row=29, end_column=c2)
        top = ws.cell(27, c1, f"{score}  {level}")
        top.font = font(12, bold=True)
        top.alignment = align("center")
        body = ws.cell(28, c1, meaning)
        body.font = font(10)
        body.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        for r in range(27, 30):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).fill = fill(RATING_FILLS[score])
                ws.cell(r, c).border = THIN
    ws.row_dimensions[29].height = 32

    ws.merge_cells("A31:L32")
    ws["A31"] = (
        "Microsoft 365 / Excel 2021+ is recommended (FILTER, MAXIFS, TEXTJOIN). "
        "A hidden Data/Calc pair powers the dashboard — you do not need to edit those sheets."
    )
    ws["A31"].font = font(10, italic=True, color=MUTED)
    ws["A31"].alignment = align("left", wrap=True)

    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TEAL
    ws.freeze_panes = "A2"
    apply_print(ws, landscape=True)


def main() -> None:
    wb = Workbook()
    wb.remove(wb.active)

    build_how_to(wb)
    build_matrix(wb)
    build_ratings(wb)
    wb.create_sheet(SHEET_CALC)
    build_dashboard(wb)
    build_employees(wb)
    build_skills(wb)
    build_years(wb)
    build_lists(wb)
    build_skill_sets(wb)
    build_settings(wb)
    build_data(wb)
    build_calc(wb)

    order = [
        SHEET_HOW,
        SHEET_MATRIX,
        SHEET_RATINGS,
        SHEET_DASH,
        SHEET_EMP,
        SHEET_SKILLS,
        SHEET_YEARS,
        SHEET_LISTS,
        SHEET_SETS,
        SHEET_SETTINGS,
        SHEET_DATA,
        SHEET_CALC,
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.properties.title = "Employee Skill Matrix"
    wb.properties.creator = "Skill Matrix"
    wb.properties.description = (
        "Rate skills 1–5 across multiple years, add years/skills/people later, "
        "and explore a person or a job profile on the dashboard."
    )

    path = "Employee_Skill_Matrix.xlsx"
    wb.save(path)
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
