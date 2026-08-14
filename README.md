# Employee Skill Matrix

An Excel workbook for a **live-event vision operating team**. Rate people from **1 (beginner) to 5 (expert)**. **Skill Matrix** is this year’s crew grid. **Dashboard** is where you compare any two years (including year 1 vs year 5).

Open **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)** in Microsoft 365 / Excel 2021 or later. Sample people and scores (2023–2026) are included. There are **no macros**.

## Skillsets

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
| **Skill Matrix** | Whole crew, **one year** (defaults to this calendar year) |
| **Ratings** | Type 1–5 for any year |
| **Dashboard** | Compare From year vs To year: team, one person, who is strongest |
| **Employees** | People. Add a person in the next yellow row |
| **Skills** | Skillset catalog. Add a skillset in the next yellow row |
| **Years** | Year list. Extra years go in a yellow row |
| **Lists** | Departments |
| **Skill Sets** | Skillset names for Dashboard ranking |
| **Settings** | Target rating (default 3) |

## Rate a skillset

1. Open **Ratings** and type **1–5** in the column for that year.
2. Open **Skill Matrix** to see the crew for the year shown (usually this calendar year).
3. Open **Dashboard**, pick **From year** and **To year**, and read the change.

## Add a person or a skillset

| Add | Where | Then |
| --- | --- | --- |
| Person | Next yellow row on **Employees** | Pick department; add a hire date |
| Skillset (to rate) | Next yellow row on **Skills** | Rate it on Ratings |
| Skillset (Dashboard ranking) | Yellow **Skill set name** on **Skill Sets**, plus a matching row | Pick it on the Dashboard |
| Department | Next yellow row on **Lists** | It appears in the Employees dropdown |
| Year | Next yellow row on **Years** (2023–2034 already exist) | Rate that column on Ratings |

## Dashboard

Yellow cells at the top:

- **From year / To year** — any two years (sample: 2023 vs 2026).
- **Department** — limits the team comparison, or **(All departments)**.
- **Employee** — that person’s From vs To scores, plus a line of their average by year.
- **Skillset** — who is strongest in the To year, with change vs From.

## Rating scale

| Rating | Level | Meaning |
| ---: | --- | --- |
| 1 | Beginner | Little or no hands-on experience |
| 2 | Basic | Can complete simple tasks with help |
| 3 | Competent | Works independently on standard tasks (default target) |
| 4 | Proficient | Handles complex work and can coach others |
| 5 | Expert | Go-to specialist; sets the standard |

## Rebuild the file

```bash
pip install -r requirements.txt
python3 create_skill_matrix.py
```
