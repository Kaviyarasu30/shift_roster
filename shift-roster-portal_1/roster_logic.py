"""
roster_logic.py
----------------
This file contains the "brain" of the roster generator.
It has NO web code in it at all — it just takes plain Python data
(team info, employee names, each person's fixed shift, rules, leaves)
and works out their day-by-day calendar for a chosen date range.

DATE RANGE, NOT JUST "A MONTH"
-------------------------------
You pick a "From date" and a "To date" — could be 2 weeks, 3 weeks,
a full month, or a range that crosses into the next month
(e.g. 9 June to 9 July). The generator doesn't care about calendar
month boundaries at all; it just walks day by day from your start
date to your end date.

FIXED SHIFT MODEL
------------------
Unlike a daily rotation (where a person's shift changes day to day),
here each employee is given ONE shift (Morning / Evening / Night) that
stays the same for the whole selected range. You decide who is on
which shift when you generate the roster — e.g. "Ravi was on Morning
last time, put him on Night this time."

The only things that change day-to-day per employee are:
  - their weekly off day(s) (rotating, fixed, or fixed-per-employee)
  - leave days (from the leave table)
Every other day, they simply work their assigned shift.
"""

from datetime import date, timedelta

MORNING = "Morning"
EVENING = "Evening"
NIGHT = "Night"
OFF = "OFF"       # weekly off
LEAVE = "LEAVE"   # employee is on approved leave

SHIFT_ORDER = [MORNING, EVENING, NIGHT]


def daterange_for_period(start_date: date, end_date: date):
    """Return every date.date() object from start_date to end_date, inclusive."""
    if end_date < start_date:
        start_date, end_date = end_date, start_date
    num_days = (end_date - start_date).days + 1
    return [start_date + timedelta(days=i) for i in range(num_days)]


def is_weekend(d: date) -> bool:
    """Saturday=5, Sunday=6 in Python's weekday() numbering."""
    return d.weekday() in (5, 6)


def build_leave_lookup(leaves, employees):
    """
    Turn the leave table (list of {employee, from, to}) into a fast lookup:
    { employee_name: set_of_date_objects_on_leave }
    """
    lookup = {emp: set() for emp in employees}
    for entry in leaves:
        emp = entry.get("employee")
        if emp not in lookup:
            continue
        start = entry.get("from")
        end = entry.get("to")
        if not start or not end:
            continue
        current = start
        while current <= end:
            lookup[emp].add(current)
            current += timedelta(days=1)
    return lookup


def assign_weekly_off_days(employees, days, rules):
    """
    Work out each employee's weekly off day(s). Three modes, controlled
    by rules["weekly_off_mode"]:

    - "fixed_per_employee" (your team's usual pattern — e.g. Keerthi is
      always off Sat & Sun, Harshita is always off Wed & Thu):
      each employee has THEIR OWN fixed weekday(s) off, the same every
      week, all month. Comes from rules["per_employee_off_days"] —
      a dict of { employee_name: [weekday numbers] }, Monday=0 ... Sunday=6.

    - "fixed" (everyone off the same days, e.g. "everyone is off every
      Saturday & Sunday"): the SAME weekday(s) are off, every week, all
      month, for every employee. Use rules["weekly_off_days"].

    - "rotating" (default): each employee gets rules["weekly_off_count"]
      off days per week, but WHICH weekday(s) those are shifts around
      week to week and person to person, so no one is always stuck
      resting on the same day.

    Returns: { employee_name: set_of_date_objects_that_are_off }
    """
    off_days = {emp: set() for emp in employees}
    if not rules.get("weekly_off_enabled", True):
        return off_days

    mode = rules.get("weekly_off_mode", "rotating")

    if mode == "fixed_per_employee":
        per_employee = rules.get("per_employee_off_days", {})
        for emp in employees:
            emp_weekdays = set(per_employee.get(emp, [5, 6]))  # default Sat+Sun if unset
            for d in days:
                if d.weekday() in emp_weekdays:
                    off_days[emp].add(d)
        return off_days

    if mode == "fixed":
        fixed_weekdays = set(rules.get("weekly_off_days", [5, 6]))  # default Sat+Sun
        for d in days:
            if d.weekday() in fixed_weekdays:
                for emp in employees:
                    off_days[emp].add(d)
        return off_days

    # ---- rotating mode (original behaviour) ----
    offs_per_week = int(rules.get("weekly_off_count", 1))

    # Split the month's days into calendar weeks (Mon-Sun chunks).
    weeks = []
    current_week = []
    for d in days:
        current_week.append(d)
        if d.weekday() == 6 or d == days[-1]:  # Sunday, or last day of month
            weeks.append(current_week)
            current_week = []

    for week_index, week in enumerate(weeks):
        for emp_index, emp in enumerate(employees):
            base = (week_index + emp_index) % len(week)
            count = min(offs_per_week, len(week))  # can't give more offs than days that exist
            for step in range(count):
                pick = (base + step) % len(week)
                off_days[emp].add(week[pick])

    return off_days


def generate_roster(team_name, start_date, end_date, employees, employee_shifts, rules, leaves):
    """
    Main entry point.

    start_date, end_date: date objects — the roster covers every day
        from start_date to end_date, inclusive (can span multiple
        calendar months, e.g. 9 June to 9 July).

    employee_shifts: { employee_name: "Morning"/"Evening"/"Night" }
        The ONE shift each employee is on for this whole date range.

    Returns a dict:
        {
          "days": [list of date objects],
          "assignments": { employee: { date: "Morning"/"Evening"/"Night"/"OFF"/"LEAVE" } },
          "warnings": [ list of text warnings, e.g. understaffed days ]
        }
    """
    days = daterange_for_period(start_date, end_date)
    leave_lookup = build_leave_lookup(leaves, employees)
    off_days = assign_weekly_off_days(employees, days, rules)

    min_staff = {
        MORNING: int(rules.get("min_staff_morning", 1)),
        EVENING: int(rules.get("min_staff_evening", 1)),
        NIGHT: int(rules.get("min_staff_night", 1)),
    }

    assignments = {emp: {} for emp in employees}
    warnings = []

    for d in days:
        working_today = {MORNING: 0, EVENING: 0, NIGHT: 0}

        for emp in employees:
            if d in leave_lookup[emp]:
                assignments[emp][d] = LEAVE
            elif d in off_days[emp]:
                assignments[emp][d] = OFF
            else:
                shift = employee_shifts.get(emp, MORNING)
                assignments[emp][d] = shift
                working_today[shift] += 1

        # Check whether each shift still meets the minimum staff rule
        # once leaves and weekly-offs are taken into account.
        for shift in SHIFT_ORDER:
            needed = min_staff[shift]
            have = working_today[shift]
            if have < needed:
                warnings.append(
                    f"{d.isoformat()} ({shift}): needed {needed} staff, "
                    f"only {have} available (others are on leave/weekly off). "
                    f"Consider adjusting who's off that day, or assigning more "
                    f"people to {shift}."
                )

    return {"days": days, "assignments": assignments, "warnings": warnings}
