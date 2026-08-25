# Skill Matrix

A browser app for a **live-event vision crew**. Rate people from **0 (not trained) to 5 (expert)**. Scores stay on this computer. You do **not** need Excel or Python.

## Run it

```bash
cd app
npm install
npm run dev
```

Open the URL Vite prints (usually `http://localhost:5173`). Use Chrome, Edge, or another modern browser.

## What you can do

| Page | What it is for |
| --- | --- |
| **How to Use** | Walkthrough and the 0–5 scale |
| **Skill Matrix** | Crew grid **by year**. Pick a year tab, type **0–5**. Blank means not scored. **0** means not trained |
| **Dashboard** | Look at Team, Skillsets, or Individual |
| **Employees** | Add or edit people. Inactive people stay in history but leave the live matrix |
| **Skills** | The eleven vision skillsets (add more if you need them) |
| **Years** | Year list. Each year is its own score grid |
| **Departments** | Vision, Sound, Lighting, Staging (add more if you need them) |
| **Settings** | Target rating (default 3), logo URL, JSON backup |

## Dashboard

Start with **Look at**, then **From year** and **To year**:

| Look at | What you see |
| --- | --- |
| **Team** | Year-on-year overall ratings (optionally one department) |
| **Skillsets** | Pick a skill. The crew is listed, strongest first. Unrated people sit at the bottom. Top row is the **best candidate**. A real **0** ranks above a blank |
| **Individual** | Pick an employee. That person vs themselves on each skill |

Overall averages include **0** and skip blanks.

## Backup

Settings → **Export JSON backup**. Import the same file on another computer. Everything lives in this browser until you export.

## Sample crew

Eight sample people and scores for 2023–2026 are included. Overwrite them with your crew.

Skillsets: Broadcast Camera Operation, PTZ Cameras Operation, Shading / CCU, ATEM Vision Switching, Barco Vision Switching, Camera Switching, Content Operation, Systems Tech, Technical Setup, Live Streaming, Troubleshooting.

## Excel workbook (optional)

The original **[Employee_Skill_Matrix.xlsx](Employee_Skill_Matrix.xlsx)** is still in this repo if you need it. New work happens in the app.

Python is only for people who maintain the Excel generator scripts. You do not need it to run the app.
