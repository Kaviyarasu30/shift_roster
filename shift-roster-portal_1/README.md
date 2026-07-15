# Shift Roster Generator — Beginner's Guide

This is a small website you run on your own computer. You fill a form
(team name, employees, rules, leaves) and click **Generate Roster** —
it creates and downloads a formatted Excel file with everyone's shifts
for the month.

This guide assumes you have **never coded before**. Follow it top to bottom.

---

## 1. The tools you'll be using (in plain English)

| Tool | What it actually is | Why we need it |
|---|---|---|
| **Python** | A programming language. Think of it as the "engine" that runs the instructions in this project. | Everything here is written in Python. |
| **pip** | Comes bundled with Python. It downloads and installs extra bits of ready-made code ("packages") that other programmers wrote, so we don't reinvent the wheel. | Used to install Flask and openpyxl (below). |
| **Flask** | A small Python package that turns a Python script into a website your browser can talk to. | Powers the form page and handles the "Generate" button. |
| **openpyxl** | A Python package that creates real `.xlsx` Excel files. | Builds the roster file you download. |
| **HTML / CSS** | The language browsers understand for page structure (HTML) and appearance (CSS). | The form you see and fill in. |
| **A code editor** (recommend **VS Code**, it's free) | A text editor built for code, with helpful colour-highlighting. | Where you'll open this project's files if you want to look at or tweak them. |
| **Terminal / Command Prompt** | A text-based way to give your computer commands, instead of clicking icons. | Used to install packages and start the website. |

You do **not** need to know how to code to use this — you just need to
follow the steps below once to get it running.

---

## 2. One-time setup

### Step 1 — Install Python
1. Go to **https://www.python.org/downloads/**
2. Download the latest version for your operating system (Windows/Mac).
3. Run the installer. **On Windows, tick the box that says "Add Python to PATH"** before clicking Install — this step trips people up if skipped.
4. Confirm it worked: open your Terminal (Mac: "Terminal" app, Windows: "Command Prompt") and type:
   ```
   python3 --version
   ```
   (On Windows it may just be `python --version`.) You should see something like `Python 3.12.3`.

### Step 2 — Get the project files
Unzip the `shift-roster-portal.zip` folder you were given, anywhere convenient (e.g. your Desktop).

### Step 3 — Open a terminal inside the project folder
- **Mac**: open Terminal, type `cd ` (with a space), then drag the `shift-roster-portal` folder into the terminal window, press Enter.
- **Windows**: open the folder in File Explorer, click the address bar, type `cmd`, press Enter.

### Step 4 — Install the two packages this project needs
In that same terminal, run:
```
pip install -r requirements.txt
```
This reads `requirements.txt` and installs Flask + openpyxl automatically. You only need to do this once (or again if you move to a new computer).

> If `pip` isn't recognized, try `pip3` instead of `pip`, and `python3` instead of `python` throughout this guide.

### Step 5 — Start the website
Still in that terminal:
```
python3 app.py
```
You'll see some text ending in something like:
```
Running on http://127.0.0.1:5000
```
That means it worked — your computer is now running the website, but only for you (it's not on the public internet).

### Step 6 — Open it in your browser
Open Chrome/Edge/Safari and go to:
```
http://127.0.0.1:5000
```
You should see the Shift Roster form.

To stop the site later, go back to the terminal and press `Ctrl + C`.

---

## 3. Using the portal

This tool uses **monthly shift rotation**: each employee is on ONE
shift (Morning, Evening, or Night) for the *entire* month, not a
different shift every day. You decide who's on which shift when you
generate the roster — next month, you just move whoever needs to
rotate to a new shift.

1. **Team Information** — team name, plus a **"From date"** and **"To date"**. Pick any range you like: 2 weeks, 3 weeks, a full month, or a range that crosses into the next month (e.g. 9 June to 9 July) — the roster covers every day from "From date" to "To date", inclusive.
2. **Employees & this month's shift** — click "+ Add employee" for each person, type their name, and pick their shift for this month from the dropdown next to it.
   - Example: if Ravi was on Morning last month and should move to Night this month, just select "Night" in his dropdown.
   - The **"🔄 Rotate everyone forward one step"** button is a shortcut: it moves every dropdown Morning → Evening, Evening → Night, Night → Morning in one click, if your whole team rotates together. You can still fix individual people by hand afterwards.
3. **Shift Details** — shown for reference only (Morning/Evening/Night timings are fixed in this version).
4. **Rules**:
   - **Weekly off** — every employee gets rest day(s) each week. In the **Rules** card, pick a style:
     - **Rotating**: set "Days off per week" (e.g. 2). Which weekday(s) each person rests on shifts around week to week and person to person, so no one is stuck resting the same day every time.
     - **Fixed — same days for everyone**: tick the exact weekday(s) that are off (e.g. Sat + Sun). Those same days apply to everyone, every single week, for the whole month.
     - **Fixed — different days per employee** (this is what most real teams use): each employee has their own permanent off day(s) — e.g. Keerthi is always off Sat & Sun, Harshita is always off Wed & Thu — staggered so the whole team isn't resting on the same day. Pick this option, then scroll up to **section 2 (Employees)** — a small Mon–Sun checklist appears under each employee's name/shift row. Tick that person's day(s) there. It stays the same every week, all month; next month you can leave it as-is or change individual people.
   - **Minimum staff per shift** — how many people must be working Morning/Evening/Night every single day. There are two kinds of checks, and **both block the download by default**:
     - **Structural check**: if you simply haven't assigned enough employees to a shift to ever meet the minimum (e.g. you require 2 on Evening but only 1 employee is on Evening at all), the tool stops immediately — no amount of adjusting off-days could fix this, so it tells you exactly which shift is short and by how much.
     - **Day-by-day check**: even with enough people overall, a leave can land on the same day as someone's weekly off, dropping that one day below the minimum. By default, Generate Roster stops here too and lists every exact date/shift that's short, so nothing short-staffed ships without you noticing.
     - **Need the file anyway?** Tick **"Generate anyway even if some days are short-staffed"** at the bottom of the Rules section — the file will download regardless, with those specific dates listed in a **Warnings** tab instead of blocking you. Useful for genuine one-off emergencies you can't immediately fix.
5. **Leaves** — for anyone with pre-approved leave, add a row with their name and the date range. They'll be marked `LEAVE` on those days and won't be scheduled.
6. Click **Generate Roster**. Your browser will download an Excel file named like `Support_Team_Roster_07_2026.xlsx`.

Open it in Excel or Google Sheets — you'll see one row per employee, one
column per day, color-coded shifts, and a legend. If the rules couldn't
be fully satisfied on some day (e.g. not enough people to meet the
minimum staff), a second sheet called **Warnings** lists exactly which
day/shift was short, so you know where to add more people or relax a rule.

---

## 4. How it actually works, in plain terms

```
Your browser (the form)
        │  (you fill it in and click Generate)
        ▼
app.py — reads what you typed, including each employee's chosen
         shift for this month
        │
        ▼
roster_logic.py — the scheduling brain:
   for every day of the month, for every employee, it:
     - marks them LEAVE if that day falls in their leave range
     - marks them OFF if it's their (rotating) weekly off day
     - otherwise, gives them the ONE shift you assigned them for
       the whole month
   then checks each day: did every shift still meet its minimum
   staff count once leaves/offs are subtracted? If not, it's logged
   as a warning.
        │
        ▼
app.py — takes that schedule and builds a color-coded Excel file
        │
        ▼
Your browser — downloads the file
```

Because each person's shift is fixed for the month, the only thing
the tool has to work out day-to-day is *who's off or on leave when* —
it's not trying to re-shuffle shifts. The Warnings sheet tells you
exactly which day/shift dropped below your minimum staffing (usually
because a leave and a weekly-off overlapped), so a human can make the
final call — e.g. shift someone's weekly off by a day.

---

## 5. If you want to change anything

- **Change shift timings**: edit the text inside `templates/index.html`, search for "Morning · 6:30 AM".
- **Change colors in the Excel file**: edit `SHIFT_COLORS` near the top of `app.py`.
- **Change the scheduling behaviour**: everything lives in `roster_logic.py`, it's the only file with actual scheduling decisions in it.

Whenever you save a change while `python3 app.py` is running, refresh
your browser — Flask automatically reloads the updated code.

---

## 6. Putting this on the internet (optional, later)

Right now this only runs on your own computer (`127.0.0.1` means
"this computer only"). If you later want coworkers to use it from
their own browsers without installing anything, you'd deploy it to a
hosting service such as **Render.com** or **PythonAnywhere.com** —
both have free tiers and accept Flask apps with barely any changes.
That's a good "next step" once you're comfortable with the local
version — happy to walk through it whenever you're ready.
