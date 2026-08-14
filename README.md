# Employee Skill Matrix

An Excel workbook for a **live-event vision operating team**. Rate people from **1 (beginner) to 5 (expert)** across **multiple years**, then use the dashboard to look at one operator or to see who is strongest for a seat.

Open **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)** in Microsoft 365 / Excel 2021 or later. Sample people and scores (2023–2026) are included — replace them with your crew.

## Vision skill sets

These are both the skills you rate and the job profiles on the Dashboard:

| Skill set | What it covers |
| --- | --- |
| **Barco Vision Switching** | Barco / LED / program switching |
| **Camera Switching** | Cutting cameras and the live picture |
| **Content Operation** | Media servers, playback, graphics |
| **Systems Tech** | Signal flow, routing, show systems |
| **Technical Setup** | Rig, patch, and line-check before doors |
| **Live Streaming** | Encode, monitor, and deliver the stream |
| **Troubleshooting** | Fixing issues under show pressure |

Each profile also expects **Technical Setup** and/or **Troubleshooting** at a minimum, so the ranking is “who can actually do that seat,” not only the one headline skill.

## Sheets

| Sheet | What it is for |
| --- | --- |
| **How to Use** | Walkthrough and the 1–5 scale |
| **Skill Matrix** | Ratings: one row per person per skill, one column per year |
| **Dashboard** | Filters, year-on-year charts, employee history, who can do a job |
| **Employees** | People. Department and job title are dropdowns. Hire date drives “since they joined” |
| **Skills** | Skill catalog. Add a row to add a skill |
| **Years** | Year list. Current calendar year is detected; extra years go in a yellow row |
| **Lists** | Departments and job titles used by the dropdowns — add items here |
| **Skill Sets** | Job profiles (a named mix of skills) for the “who can do the job” ranking |
| **Settings** | Target rating (default 3) |

## Rate a skill

1. Open **Skill Matrix**.
2. Find the person and skill (use the Employee filter on the header if you want one person at a time).
3. Enter **1–5** in the column for that year. The current calendar year is highlighted.
4. **First**, **Latest**, **vs last year**, and **Since joined** fill in on their own.

## Add a new year (automatic each year)

Years **2023–2034** are already columns. When a new year starts, rate that column — you do not copy last year’s scores. Charts pick up a year as soon as it has ratings.

On **Years** (or **Add / manage years** on the Dashboard):

- If this calendar year is already in the list, a button takes you to the Skill Matrix to rate it.
- If it is missing, the button jumps to the next yellow **Year** row so you can type it (for example **2035**). Keep years in order, oldest at the top.

## Add a skill, person, department, or job title

| Add | Where | Then |
| --- | --- | --- |
| Skill | Next yellow row on **Skills** | Rate it on the Skill Matrix |
| Person | Next yellow row on **Employees** | Pick department and title from the dropdowns; add a hire date |
| Department or job title | Next yellow row on **Lists** | It appears in the Employees dropdowns. Add it on Lists first |
| Skill set (job profile) | **Skill Sets** — same set name, one skill per row, optional minimum 1–5 | Pick it on the Dashboard to rank people |

## Dashboard filters

Yellow cells at the top:

- **Employee** — that person’s skills, last year vs latest, and change since they joined (first recorded rating after hire).
- **Skill set** — who currently looks strongest for that job (latest-year average on the required skills, vs last year, and how many skills meet the minimum).
- **Department** — limits the team year-on-year chart and skill table, or leave **(All departments)**.

Line charts show **year on year on year**, not only two years. Years with no ratings stay off the chart until you enter scores.

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
