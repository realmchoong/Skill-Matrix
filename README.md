# Employee Skill Matrix

An Excel workbook for a **live-event vision operating team**. Rate people from **1 (beginner) to 5 (expert)**. **Skill Matrix** is this year’s crew grid. **Dashboard** starts with **Look at**: Team, Skillsets, or Individual.

Open **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)** in Microsoft 365 / Excel 2021 or later. Sample people and scores (2023–2026) are included. There are **no macros**.

## Dashboard

Start with **Look at**, then **From year** and **To year**:

| Look at | What you see |
| --- | --- |
| **Team** | Year-on-year **overall** ratings (optionally one department: Vision, Sound, Lighting, Staging) |
| **Skillsets** | Pick a **Skillset** (the job to fill). Everyone is ranked on that skill. Top row is the **best candidate** |
| **Individual** | Pick an **Employee**. That person vs themselves on each skillset |

The yellow **Skillset** dropdown is always there. It is used when Look at is **Skillsets**.

## Skillsets

Broadcast Camera Operation, PTZ Cameras Operation, Shading / CCU, ATEM Vision Switching, Barco Vision Switching, Camera Switching, Content Operation, Systems Tech, Technical Setup, Live Streaming.

Add more: yellow row on **Skills**. The new name appears on Skill Matrix and in the Dashboard Skillsets list.

## Sheets

| Sheet | What it is for |
| --- | --- |
| **How to Use** | Walkthrough and the 1–5 scale |
| **Skill Matrix** | Whole crew, **one year** (this calendar year) |
| **Dashboard** | Look at Team, Skillsets, or Individual |
| **Employees** | Add a person in the next yellow row |
| **Skills** | Skillset catalog (ratings and Dashboard ranking) |
| **Years** | Year list |
| **Lists** | Departments |
| **Settings** | Target rating (default 3) |
| **Ratings** (hidden) | Type 1–5 for any year. Recall: right-click a tab → **Unhide** → Ratings. Hide it again when done. |

## Rate a skillset

1. Right-click a sheet tab → **Unhide** → **Ratings**. Type **1–5** in the year column.
2. Right-click the Ratings tab → **Hide**.
3. Open **Skill Matrix** for this year’s crew grid.
4. Open **Dashboard**, choose **Look at**, then From / To years.

## Rebuild the file

```bash
pip install -r requirements.txt
python3 create_skill_matrix.py
```
