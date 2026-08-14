# Employee Skill Matrix

An Excel workbook for rating employee skills from **1 (beginner) to 5 (expert)**, adding new skills over time, and comparing **last year’s ratings with this year’s**.

Open **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)** in Excel (or Google Sheets). Sample people and scores are included so you can see the layout — replace them with your team.

## Sheets

| Sheet | What it is for |
| --- | --- |
| **How to Use** | Short walkthrough and the 1–5 scale |
| **Skill Matrix** | Where you enter ratings. Each skill has Last year, This year, and Change |
| **Dashboard** | Team, skill, and employee averages, plus who improved or declined |
| **Employees** | Names and roles. Add a row here to add a person to the matrix |
| **Skills** | Skill catalog. Add a row here to add a skill to the matrix |
| **Settings** | This year, last year, and the target rating (default 3) |

## Rate a skill

1. Open **Skill Matrix**.
2. Find the person (rows) and skill (column group).
3. Enter a whole number **1–5** in Last year and/or This year.
4. **Change** fills in on its own (this year minus last year).

Cells color themselves: 1 red → 5 green. Green change means improved; red means declined.

## Add a skill later

1. Open **Skills**.
2. In the next yellow row, type a **Skill Name** (category and description are optional).
3. Return to **Skill Matrix** — the next empty column group shows that skill.
4. Enter last-year and this-year ratings for each person.

Sixteen skill slots are already set up. Unused slots stay blank until you name them.

## Add an employee later

1. Open **Employees**.
2. In the next yellow row, keep the Employee ID and fill in **Full Name** and **Department**.
3. The Skill Matrix adds a row. Rate each skill.

## Compare last year with this year

- On **Skill Matrix**, read the **Change** column beside each skill.
- On **Dashboard**, see team average, each skill’s average, each person’s average, skills below the target, and how many 1s–5s you gave.
- On **Settings**, change **This year** / **Last year** (the column headers follow) and **Target rating** (scores below this are flagged).

### Starting a new review year

1. Copy every This Year rating and paste **values** into Last Year.
2. Clear This Year and enter the new scores.
3. Update the years on **Settings**.

## Rating scale

| Rating | Level | Meaning |
| ---: | --- | --- |
| 1 | Beginner | Little or no hands-on experience |
| 2 | Basic | Can complete simple tasks with help |
| 3 | Competent | Works independently on standard tasks (default target) |
| 4 | Proficient | Handles complex work and can coach others |
| 5 | Expert | Go-to specialist; sets the standard |

## Rebuild the file

If you change the generator and want a fresh workbook (this replaces sample data):

```bash
pip install -r requirements.txt
python3 create_skill_matrix.py
```
