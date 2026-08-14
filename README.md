# Employee Skill Matrix

An Excel workbook for a **live-event vision operating team**. Rate people from **1 (beginner) to 5 (expert)** across **multiple years**, then compare any two years (including year 1 vs year 5) and see who is strongest for a skillset.

Open **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)** in Microsoft 365 / Excel 2021 or later. Sample people and scores (2023–2026) are included — replace them with your crew. There are **no macros**.

## Skillsets

These are both the skillsets you rate and the Dashboard ranking list:

| Skillset | Category |
| --- | --- |
| Broadcast Camera Operation | Cameras |
| PTZ Cameras Operation | Cameras |
| Shading / CCU | Cameras |
| ATEM Vision Switching | Switching |
| Barco Vision Switching | Switching |
| Camera Switching | Switching |
| Content Operation | Content |
| Systems Tech | Systems |
| Technical Setup | Systems |
| Live Streaming | Streaming |

Add more later: yellow row on **Skills**, then add the name on **Skill Sets** so the Dashboard can rank it.

## Departments

Vision, Sound, Lighting, Staging. Add another on **Lists**, then pick it on **Employees**. There is no job title field.

## Sheets

| Sheet | What it is for |
| --- | --- |
| **How to Use** | Walkthrough and the 1–5 scale |
| **Skill Matrix** | Whole-crew glance: pick **From year** and **To year** (any two years) |
| **Ratings** | Type 1–5 for any year (source of truth) |
| **Dashboard** | Filters, year-on-year charts, employee history, who can do a skillset |
| **Employees** | People. Add a person in the next yellow row |
| **Skills** | Skillset catalog. Add a skillset in the next yellow row |
| **Years** | Year list. Current calendar year is detected; extra years go in a yellow row |
| **Lists** | Departments used by the Employees dropdown |
| **Skill Sets** | Skillset names for Dashboard ranking. Add a name in the yellow column |
| **Settings** | Target rating (default 3) |

## Rate a skillset

1. Open **Ratings** and type **1–5** in the column for that year (every person × skillset).
2. Open **Skill Matrix** to see the **whole crew at once**. Each skillset shows two years and the change.
3. At the top, pick **From year** and **To year** (for example `2023` and `2026`). Every employee’s columns update together.

## Add a person or a skillset

| Add | Where | Then |
| --- | --- | --- |
| Person | Next yellow row on **Employees** | Pick department; add a hire date |
| Skillset (to rate) | Next yellow row on **Skills** | Rate it on Ratings / Skill Matrix |
| Skillset (Dashboard ranking) | Yellow **Skill set name** on **Skill Sets**, plus a matching row (same name + skill) | Pick it on the Dashboard |
| Department | Next yellow row on **Lists** | It appears in the Employees dropdown |
| Year | Next yellow row on **Years** (2023–2034 already exist) | Rate that column on Ratings |

## Dashboard filters

Yellow cells at the top:

- **Employee** — that person’s skillsets, last year vs latest, and change since they joined.
- **Skillset** — who currently looks strongest for that job.
- **Department** — limits the team year-on-year chart, or leave **(All departments)**.

Line charts show **year on year**, not only two years.

## Rating scale

| Rating | Level | Meaning |
| ---: | --- | --- |
| 1 | Beginner | Little or no hands-on experience |
| 2 | Basic | Can complete simple tasks with help |
| 3 | Competent | Works independently on standard tasks (default target) |
| 4 | Proficient | Handles complex work and can coach others |
| 5 | Expert | Go-to specialist; sets the standard |

## Rebuild the file

This replaces sample data:

```bash
pip install -r requirements.txt
python3 create_skill_matrix.py
```
