# Employee Skill Matrix

An Excel workbook for a **live-event vision operating team**. Rate people from **1 (beginner) to 5 (expert)**. **Skill Matrix** is this year’s crew grid. **Dashboard** starts with **Look at**: Team, Skillsets, or Individual.

**You only need Excel.** Do not install Python. There are **no macros**.

1. Download **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)**.
2. Open it in **Microsoft 365** or **Excel 2021** or later.
3. Add people, skillsets, and scores in the sheets below. Keep using this same file.

Sample people and scores (2023–2026) are included. You can overwrite them.

## Dashboard

Start with **Look at**, then **From year** and **To year**:

| Look at | What you see |
| --- | --- |
| **Team** | Year-on-year **overall** ratings (optionally one department: Vision, Sound, Lighting, Staging) |
| **Skillsets** | Pick a **Skillset** (the job to fill). The whole crew is listed, strongest first. People not yet rated for the To year sit at the bottom. Top row is the **best candidate** |
| **Individual** | Pick an **Employee**. That person vs themselves on each skillset |

The yellow **Skillset** dropdown is always there. It is used when Look at is **Skillsets**.

The company logo is inserted **once**. On **How to Use**, click the yellow top-left box, then **Insert → Pictures → Place in Cell**. If the picture sits on top of the grid, right-click it → **Place in Cell**. The other tabs show that same cell.

Alternatively, paste a **https** logo link in **Settings → Logo URL** (OneDrive or the web). That also appears on every tab.

## Skillsets

Broadcast Camera Operation, PTZ Cameras Operation, Shading / CCU, ATEM Vision Switching, Barco Vision Switching, Camera Switching, Content Operation, Systems Tech, Technical Setup, Live Streaming, Troubleshooting.

Add more: yellow row on **Skills**. The new name appears on Skill Matrix and in the Dashboard Skillsets list.

## Sheets

| Sheet | What it is for |
| --- | --- |
| **How to Use** | Walkthrough and the 1–5 scale |
| **Skill Matrix** | Whole crew, **this year**. Type **1–5** here. Scores also appear on Ratings |
| **Dashboard** | Look at Team, Skillsets, or Individual |
| **Employees** | Add a person in the next yellow row (fill clears when you type) |
| **Skills** | Skillset catalog (ratings and Dashboard ranking) |
| **Years** | Year list |
| **Lists** | Departments |
| **Settings** | Target rating (default 3) |
| **Ratings** (hidden) | All years. This year follows Skill Matrix. Type older years here. Unhide: right-click a tab → **Unhide** → Ratings |

## Rate a skillset

1. Open **Skill Matrix**. Type **1–5** in that person’s skill cell. Empty cells are yellow until you type.
2. Unhide **Ratings** if you want to check: this year’s column matches the grid. Hide it again when done.
3. To score an **older year**, Unhide **Ratings** and type in that year column.
4. Open **Dashboard**, choose **Look at**, then From / To years.

## Update in Excel (no Python)

Keep using **this same workbook**. Type values in the sheets below. Do **not** paste Python. Do **not** type over name/overall formulas on **Skill Matrix**, or over the Dashboard table, **Data**, or **Calc**.

| What you want | Where to type |
| --- | --- |
| Add a person | **Employees** — next yellow row: Full Name, Department, Hire Date, Status |
| Add a skillset | **Skills** — next yellow row: Skill Name (then Category / Description) |
| Add a department | **Lists** — next yellow row, then choose it on Employees |
| Add a year (2035 or later) | **Years** — next yellow Year cell, in order, oldest to newest |
| Type this year’s 1–5 | **Skill Matrix** — the person’s skill cell. It also appears on Ratings |
| Type an older year’s 1–5 | **Ratings** (Unhide) — that year column |
| Change Look at / years / filters | **Dashboard** — yellow dropdowns only |
| Target rating or logo URL | **Settings** |
| Company logo picture | **How to Use** — yellow top-left box, Insert → Pictures → **Place in Cell** |

If a cell already shows a formula in the formula bar (starts with `=`), leave it. Skill Matrix and Dashboard read from Employees, Skills, Years, and Ratings automatically.

Overwrite sample names and scores with your crew. Save the file as you go (`Ctrl+S` / `Cmd+S`).

## Get new features into an existing file

Layout changes (yellow-fill behaviour, Dashboard Skillsets list, logo box, and so on) live in a **new** `Employee_Skill_Matrix.xlsx`. They cannot be pasted as Python, and they cannot be merged into the old file automatically without Python.

Do this in Excel only:

1. Keep your current workbook. File → **Save As** a backup (for example `Skill_Matrix_backup.xlsx`).
2. Download the latest **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)** and open it. You now have two Excel windows: **Old** (your data) and **New** (the features).
3. Copy **values only** from Old into New. After you copy, right-click the destination → **Paste Special** → **Values** (not a normal paste). That keeps the new formulas intact.

Copy in this order:

| From Old | Paste Values into New | What to copy |
| --- | --- | --- |
| **Employees** | **Employees**, starting at **B6** | Full Name, Department, Hire Date, Status. Keep people in the **same row order**. Do not paste over column A (Employee ID) |
| **Skills** | **Skills**, starting at **B6** | Skill Name, Category, Description. Same row order. Do not paste over column A |
| **Lists** | **Lists**, starting at **A6** | Department names |
| **Years** | **Years**, starting at **B8** | Extra years you typed (if any), plus label/notes |
| **Ratings** (Unhide on both files) | **Ratings**, starting at **E6** | Only **older year** 1–5 columns. Do **not** paste this year’s column (it follows Skill Matrix) or columns A–D |
| **Skill Matrix** scores | **Skill Matrix**, starting at **D6** | This year’s 1–5 grid only |
| **Settings** | **Settings** **B6** and **B7** | Target rating and Logo URL |
| **Dashboard** yellow cells | Same yellow cells | Look at, From year, To year, Department, Skillset, Employee |

4. If you had a logo, insert it again on **How to Use** with **Place in Cell**.
5. Hide **Ratings** on the new file. Check **Skill Matrix** and **Dashboard**.
6. Save the **new** file and use that one from now on.

If a paste overwrites a formula (formula bar still starts with `=` on Skill Matrix / Dashboard / Ratings A–D), Undo (`Ctrl+Z`) and paste into the value cells listed above instead.

## Rebuild the file (optional)

Python is **not required** to use the workbook. It is only for people who maintain this repository and need to rebuild the `.xlsx` from the generator scripts.

You do **not** need a new file to add people or skillsets — use the yellow rows on **Employees** and **Skills** (the fill disappears when you type), and type 1–5 on **Ratings**.

`python3 create_skill_matrix.py` **replaces** `Employee_Skill_Matrix.xlsx` with a new sample workbook. Scores you typed would not be in that new file.

When you want a newer layout (Dashboard, logo box, extra skillsets, and so on) **and** you already have real data:

```bash
python3 create_skill_matrix.py --from Employee_Skill_Matrix.xlsx
```

That builds the new workbook, copies your people, skillsets, years, departments, 1–5 ratings, Settings, Dashboard picks, and any logo pictures into it, and saves a timestamped backup first.

To write a separate file instead of replacing this one:

```bash
python3 create_skill_matrix.py --from Employee_Skill_Matrix.xlsx -o Employee_Skill_Matrix_new.xlsx
```

Fresh sample file (no copy):

```bash
pip install -r requirements.txt
python3 create_skill_matrix.py
```
