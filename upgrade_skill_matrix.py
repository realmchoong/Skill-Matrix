"""Copy people, skills, ratings, and logo from an existing skill-matrix workbook."""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from create_skill_matrix import (
    MATRIX_HEADER_ROW,
    NUM_EMPLOYEE_SLOTS,
    NUM_LIST_SLOTS,
    NUM_SKILL_SLOTS,
    NUM_YEAR_SLOTS,
    NAV_SHEETS,
    SHEET_HOW,
    SHEET_DASH,
    SHEET_EMP,
    SHEET_LISTS,
    SHEET_RATINGS,
    SHEET_SETTINGS,
    SHEET_SKILLS,
    SHEET_YEARS,
    SKILLS,
    YEAR_FIRST_COL,
    matrix_row,
    year_col,
)


def copy_user_data(source_path: str | Path, dest: Workbook) -> None:
    src = load_workbook(source_path, data_only=False)
    missing = [name for name in (SHEET_EMP, SHEET_SKILLS, SHEET_RATINGS) if name not in src.sheetnames]
    if missing:
        raise SystemExit(
            f"{source_path} is missing sheet(s): {', '.join(missing)}. "
            "This upgrade only works with a Skill Matrix workbook."
        )

    employees = _read_employees(src[SHEET_EMP])
    skills = _read_skills(src[SHEET_SKILLS])
    years = _read_years(src[SHEET_YEARS]) if SHEET_YEARS in src.sheetnames else []
    departments = _read_departments(src[SHEET_LISTS]) if SHEET_LISTS in src.sheetnames else []
    ratings = _read_ratings(src, employees, skills, years)
    settings = _read_settings(src)
    dashboard = _read_dashboard(src)

    new_skills = _merge_skills(skills)
    _write_employees(dest[SHEET_EMP], employees)
    _write_skills(dest[SHEET_SKILLS], new_skills)
    if years:
        _write_years(dest[SHEET_YEARS], years)
    if departments:
        _write_departments(dest[SHEET_LISTS], departments)
    _write_ratings(dest[SHEET_RATINGS], dest[SHEET_EMP], dest[SHEET_SKILLS], dest[SHEET_YEARS], ratings)
    _write_settings(dest, settings)
    _write_dashboard(dest, dashboard)
    images = _collect_images(src)
    dest._pending_images = images  # consumed just before save

    print(
        f"Copied {len(employees)} people, {len(new_skills)} skillsets, "
        f"{len(ratings)} ratings from {source_path}"
    )
    extra = [name for name, *_rest in new_skills if name.lower() not in {s[0].lower() for s in skills}]
    if extra:
        print("Added new skillsets not in the old file: " + ", ".join(extra))
    skipped_people = max(0, _count_named_rows(src[SHEET_EMP], "Full Name") - NUM_EMPLOYEE_SLOTS)
    skipped_skills = max(0, _count_named_rows(src[SHEET_SKILLS], "Skill Name") - NUM_SKILL_SLOTS)
    if skipped_people:
        print(f"Warning: {skipped_people} extra people did not fit (limit {NUM_EMPLOYEE_SLOTS}).")
    if skipped_skills:
        print(f"Warning: {skipped_skills} extra skillsets did not fit (limit {NUM_SKILL_SLOTS}).")


def _norm(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _header_cols(ws: Worksheet, header_row: int = 5) -> dict[str, int]:
    cols: dict[str, int] = {}
    for col in range(1, 20):
        value = _norm(ws.cell(header_row, col).value)
        if value:
            cols[value] = col
    return cols


def _find_header_row(ws: Worksheet, title: str, max_row: int = 12) -> int | None:
    for row in range(1, max_row + 1):
        for col in range(1, 20):
            if _norm(ws.cell(row, col).value) == title:
                return row
    return None


def _is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _read_block(ws: Worksheet, header: str, columns: list[str], start_hint: int = 5) -> list[dict]:
    header_row = _find_header_row(ws, header) or start_hint
    cols = _header_cols(ws, header_row)
    if header not in cols:
        return []
    rows = []
    for row in range(header_row + 1, header_row + 1 + max(NUM_EMPLOYEE_SLOTS, NUM_SKILL_SLOTS, NUM_YEAR_SLOTS) + 8):
        record = {}
        empty = True
        for name in columns:
            col = cols.get(name)
            value = ws.cell(row, col).value if col else None
            if _is_formula(value):
                value = None
            record[name] = value
            if _norm(value):
                empty = False
        if empty:
            continue
        rows.append(record)
    return rows


def _count_named_rows(ws: Worksheet, header: str) -> int:
    header_row = _find_header_row(ws, header)
    if not header_row:
        return 0
    cols = _header_cols(ws, header_row)
    col = cols.get(header)
    if not col:
        return 0
    count = 0
    for row in range(header_row + 1, header_row + 80):
        if _norm(ws.cell(row, col).value) and not _is_formula(ws.cell(row, col).value):
            count += 1
    return count


def _read_employees(ws: Worksheet) -> list[tuple]:
    records = _read_block(ws, "Full Name", ["Employee ID", "Full Name", "Department", "Hire Date", "Status"])
    people = []
    for rec in records:
        name = _norm(rec.get("Full Name"))
        if not name:
            continue
        people.append(
            (
                rec.get("Employee ID") or "",
                name,
                rec.get("Department") or "",
                rec.get("Hire Date"),
                rec.get("Status") or "Active",
            )
        )
    return people[:NUM_EMPLOYEE_SLOTS]


def _read_skills(ws: Worksheet) -> list[tuple[str, str, str]]:
    records = _read_block(ws, "Skill Name", ["Skill Name", "Category", "Description"])
    skills = []
    for rec in records:
        name = _norm(rec.get("Skill Name"))
        if not name:
            continue
        skills.append((name, rec.get("Category") or "", rec.get("Description") or ""))
    return skills[:NUM_SKILL_SLOTS]


def _read_years(ws: Worksheet) -> list[tuple]:
    header_row = _find_header_row(ws, "Year") or 7
    cols = _header_cols(ws, header_row)
    year_col_i = cols.get("Year", 2)
    label_col = cols.get("Label", 3)
    notes_col = cols.get("Notes", 4)
    years = []
    for row in range(header_row + 1, header_row + 1 + NUM_YEAR_SLOTS + 4):
        raw = ws.cell(row, year_col_i).value
        if _is_formula(raw) or raw in (None, ""):
            continue
        try:
            year = int(raw)
        except (TypeError, ValueError):
            continue
        years.append((year, ws.cell(row, label_col).value or "", ws.cell(row, notes_col).value or ""))
    return years[:NUM_YEAR_SLOTS]


def _read_departments(ws: Worksheet) -> list[str]:
    header_row = _find_header_row(ws, "Department") or 5
    cols = _header_cols(ws, header_row)
    col = cols.get("Department", 1)
    depts = []
    for row in range(header_row + 1, header_row + 1 + NUM_LIST_SLOTS + 4):
        value = ws.cell(row, col).value
        if _is_formula(value) or not _norm(value):
            continue
        if _norm(value) not in depts:
            depts.append(_norm(value))
    return depts[:NUM_LIST_SLOTS]


def _read_ratings(src: Workbook, employees: list, skills: list, years: list) -> dict[tuple[str, str, int], int]:
    ws = src[SHEET_RATINGS]
    header_row = _find_header_row(ws, "Employee") or MATRIX_HEADER_ROW
    category_col = None
    first_summary = None
    for col in range(1, 40):
        title = _norm(ws.cell(header_row, col).value)
        if title == "Category":
            category_col = col
        if title in {"First", "Latest"}:
            first_summary = col
            break
    year_start = (category_col + 1) if category_col else YEAR_FIRST_COL
    year_end = (first_summary - 1) if first_summary else year_start + NUM_YEAR_SLOTS - 1

    year_headers = []
    src_years = [y[0] for y in years]
    for i, col in enumerate(range(year_start, year_end + 1)):
        header = ws.cell(header_row, col).value
        year = None
        if isinstance(header, (int, float)):
            year = int(header)
        elif i < len(src_years):
            year = src_years[i]
        year_headers.append((col, year))

    n_emp_slots, n_skill_slots = _detect_rating_slots(ws, header_row)
    ratings: dict[tuple[str, str, int], int] = {}
    emp_names = [e[1] for e in employees]
    skill_names = [s[0] for s in skills]
    data_start = header_row + 1
    for emp_i, emp_name in enumerate(emp_names):
        if emp_i >= n_emp_slots:
            break
        for skill_i, skill_name in enumerate(skill_names):
            if skill_i >= n_skill_slots:
                break
            row = data_start + emp_i * n_skill_slots + skill_i
            for col, year in year_headers:
                if year is None:
                    continue
                value = ws.cell(row, col).value
                if _is_formula(value) or value in (None, ""):
                    continue
                try:
                    rating = int(value)
                except (TypeError, ValueError):
                    continue
                if 1 <= rating <= 5:
                    ratings[(emp_name.lower(), skill_name.lower(), int(year))] = rating
    return ratings


def _detect_rating_slots(ws: Worksheet, header_row: int) -> tuple[int, int]:
    data_start = header_row + 1
    n_rows = max(0, (ws.max_row or data_start) - data_start + 1)
    for n_emp in (NUM_EMPLOYEE_SLOTS, 8, 12, 20, 24, 32):
        if n_rows and n_rows % n_emp == 0:
            n_skill = n_rows // n_emp
            if 4 <= n_skill <= 40:
                return n_emp, n_skill
    return NUM_EMPLOYEE_SLOTS, NUM_SKILL_SLOTS


def _read_settings(src: Workbook) -> dict:
    if SHEET_SETTINGS not in src.sheetnames:
        return {}
    ws = src[SHEET_SETTINGS]
    out = {}
    for row in range(1, 20):
        label = _norm(ws.cell(row, 1).value)
        value = ws.cell(row, 2).value
        if label == "Target rating" and isinstance(value, (int, float)) and not _is_formula(value):
            out["target"] = int(value)
        if label == "Logo URL" and _norm(value) and not _is_formula(value):
            out["logo_url"] = _norm(value)
    return out


def _read_dashboard(src: Workbook) -> dict:
    if SHEET_DASH not in src.sheetnames:
        return {}
    ws = src[SHEET_DASH]
    out = {}
    for key, coord in (
        ("look_at", "B5"),
        ("from_year", "D5"),
        ("to_year", "F5"),
        ("department", "B6"),
        ("skillset", "D6"),
        ("employee", "F6"),
    ):
        value = ws[coord].value
        if value in (None, "") or _is_formula(value):
            continue
        out[key] = value
    return out


def _merge_skills(old: list[tuple[str, str, str]]) -> list[tuple[str, str, str]]:
    merged = list(old)
    have = {name.lower() for name, _cat, _desc in old}
    for _sid, name, category, desc in SKILLS:
        if name.lower() not in have:
            merged.append((name, category, desc))
            have.add(name.lower())
    return merged[:NUM_SKILL_SLOTS]


def _write_employees(ws: Worksheet, employees: list[tuple]) -> None:
    for i in range(NUM_EMPLOYEE_SLOTS):
        row = 6 + i
        emp_id = f"EMP-{i + 1:03d}"
        if i < len(employees):
            _eid, name, dept, hired, status = employees[i]
            values = [emp_id, name, dept, hired, status]
        else:
            values = [emp_id, "", "", None, ""]
        for col, value in enumerate(values, 1):
            ws.cell(row, col, value)
        ws.cell(row, 4).number_format = "YYYY-MM-DD"


def _write_skills(ws: Worksheet, skills: list[tuple[str, str, str]]) -> None:
    for i in range(NUM_SKILL_SLOTS):
        row = 6 + i
        skill_id = f"SK-{i + 1:03d}"
        if i < len(skills):
            name, category, desc = skills[i]
            values = [skill_id, name, category, desc]
        else:
            values = [skill_id, "", "", ""]
        for col, value in enumerate(values, 1):
            ws.cell(row, col, value)


def _write_years(ws: Worksheet, years: list[tuple]) -> None:
    header_row = _find_header_row(ws, "Year") or 7
    for i in range(NUM_YEAR_SLOTS):
        row = header_row + 1 + i
        if i < len(years):
            year, label, notes = years[i]
            ws.cell(row, 2, year)
            ws.cell(row, 3, label)
            ws.cell(row, 4, notes)
        else:
            ws.cell(row, 2, "")
            ws.cell(row, 3, "")
            ws.cell(row, 4, "")


def _write_departments(ws: Worksheet, departments: list[str]) -> None:
    header_row = _find_header_row(ws, "Department") or 5
    for i in range(NUM_LIST_SLOTS):
        row = header_row + 1 + i
        cell = ws.cell(row, 1)
        cell.value = departments[i] if i < len(departments) else ""


def _write_ratings(
    ratings_ws: Worksheet,
    emp_ws: Worksheet,
    skill_ws: Worksheet,
    years_ws: Worksheet,
    ratings: dict[tuple[str, str, int], int],
) -> None:
    emp_names = [_norm(emp_ws.cell(6 + i, 2).value) for i in range(NUM_EMPLOYEE_SLOTS)]
    skill_names = [_norm(skill_ws.cell(6 + i, 2).value) for i in range(NUM_SKILL_SLOTS)]
    years = []
    header_row = _find_header_row(years_ws, "Year") or 7
    for i in range(NUM_YEAR_SLOTS):
        raw = years_ws.cell(header_row + 1 + i, 2).value
        try:
            years.append(int(raw))
        except (TypeError, ValueError):
            years.append(None)

    for emp_i in range(NUM_EMPLOYEE_SLOTS):
        for skill_i in range(NUM_SKILL_SLOTS):
            row = matrix_row(emp_i, skill_i)
            emp = emp_names[emp_i].lower()
            skill = skill_names[skill_i].lower()
            for slot in range(NUM_YEAR_SLOTS):
                col = year_col(slot)
                year = years[slot] if slot < len(years) else None
                value = ratings.get((emp, skill, year)) if emp and skill and year else None
                ratings_ws.cell(row, col).value = value


def _write_settings(dest: Workbook, settings: dict) -> None:
    if SHEET_SETTINGS not in dest.sheetnames:
        return
    ws = dest[SHEET_SETTINGS]
    for row in range(1, 20):
        label = _norm(ws.cell(row, 1).value)
        if label == "Target rating" and "target" in settings:
            ws.cell(row, 2).value = settings["target"]
        if label == "Logo URL" and "logo_url" in settings:
            ws.cell(row, 2).value = settings["logo_url"]


def _write_dashboard(dest: Workbook, dashboard: dict) -> None:
    if not dashboard or SHEET_DASH not in dest.sheetnames:
        return
    ws = dest[SHEET_DASH]
    mapping = {
        "look_at": "B5",
        "from_year": "D5",
        "to_year": "F5",
        "department": "B6",
        "skillset": "D6",
        "employee": "F6",
    }
    for key, coord in mapping.items():
        if key in dashboard:
            ws[coord].value = dashboard[key]


def embed_pending_images(dest: Workbook, tmp_dir: str | Path) -> int:
    """Stamp a copied logo onto every tab so a floating picture still appears everywhere."""
    jobs = getattr(dest, "_pending_images", None) or []
    dest._pending_images = []
    if not jobs:
        return 0
    blob = jobs[0][1]
    suffix = ".jpg" if blob[:2] == b"\xff\xd8" else ".png"
    path = Path(tmp_dir) / f"logo{suffix}"
    path.write_bytes(blob)
    copied = 0
    for name in NAV_SHEETS:
        if name not in dest.sheetnames:
            continue
        new_img = XLImage(str(path))
        new_img.anchor = "A1"
        dest[name].add_image(new_img)
        copied += 1
    if copied:
        print(f"Copied the logo onto {copied} tabs")
    return copied


def _images_on(ws: Worksheet) -> list[bytes]:
    blobs = []
    for img in list(getattr(ws, "_images", [])):
        try:
            blobs.append(img._data())
        except Exception:
            continue
    return blobs


def _collect_images(src: Workbook) -> list[tuple[str, bytes, object]]:
    blobs: list[bytes] = []
    if SHEET_HOW in src.sheetnames:
        blobs.extend(_images_on(src[SHEET_HOW]))
    for name in src.sheetnames:
        if name == SHEET_HOW:
            continue
        blobs.extend(_images_on(src[name]))
    if not blobs:
        return []
    return [(SHEET_HOW, blobs[0], "A1")]
