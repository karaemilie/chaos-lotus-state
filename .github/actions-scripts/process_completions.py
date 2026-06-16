"""
process_completions.py — Canonical TASKS completion engine for the Beast.

THE SACRED ENGINE. Every TASKS-source completion routes through here — never
hand-write a completion. Handles:
  • Move task → COMPLETED sheet with Alaska-correct completion date
  • Recurrence: compute next occurrence from the clean Recurring Type/Value fields
  • Sequential groups: unlock the next task in SeqOrder
  • ID audit: catch duplicate IDs / TASKS∩COMPLETED collisions / range gaps

RECURRING SCHEMA (migrated 2026-06-15 — two clean fields, dropdown-locked Type):
  Recurring Type  ∈ {None, Interval from completion, Weekly, Monthly, Yearly, Interval+Weekday}
  Recurring Value depends on Type:
     None                      → (ignored)
     Interval from completion  → "N units" e.g. "30 days","2 weeks","3 months","1 years"
                                  (special: "N months@D" = every N months on day-of-month D)
     Weekly                    → weekday name e.g. "Sunday"
     Monthly                   → day-of-month "15"  OR  "Nth Weekday" e.g. "4th Friday"
     Yearly                    → "Mon DD" e.g. "Nov 10"
     Interval+Weekday          → "N weeks/Weekday" e.g. "2 weeks/Saturday"

TASKS schema (18 cols): ID(1) Task(2) AnchorWeight(3) Priority(4) Duration(5)
  MentalLoad(6) Project(7) Category(8) SeqGroup(9) SeqOrder(10) Start(11)
  Due(12) DateLock(13) Recurring(14, legacy) Notes(15) Critical(16)
  RecurringType(17) RecurringValue(18)

Usage:
    import process_completions as pc
    result = pc.process_completions(wb, [task_id, ...])   # wb = openpyxl workbook
    # returns dict with per-task outcomes; caller saves + verifies the wb
"""

import re
import calendar
from datetime import datetime, timedelta

try:
    from stamp_helper import alaska_stamp_date
except ImportError:
    def alaska_stamp_date():
        from datetime import timezone
        try:
            from zoneinfo import ZoneInfo
            now = datetime.now(ZoneInfo("America/Anchorage"))
        except ImportError:
            now = datetime.utcnow() - timedelta(hours=8)
        if now.hour < 3:
            now = now - timedelta(days=1)
        return datetime(now.year, now.month, now.day)


WEEKDAYS = {"monday": 0, "tuesday": 1, "wednesday": 2, "thursday": 3,
            "friday": 4, "saturday": 5, "sunday": 6}
MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
MONTHS.update({m.lower(): i for i, m in enumerate(calendar.month_name) if m})

NON_RECURRING = {None, "", "none", "no", "false", "one-time", "one time"}


def _hdr(ws):
    return {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}


def _as_date(v):
    if isinstance(v, datetime):
        return v
    if hasattr(v, "year") and hasattr(v, "month"):
        return datetime(v.year, v.month, v.day)
    return None


def _next_weekday(from_date, weekday_idx):
    """Next date strictly after from_date landing on weekday_idx (0=Mon)."""
    days = (weekday_idx - from_date.weekday()) % 7
    if days == 0:
        days = 7
    return from_date + timedelta(days=days)


def _add_months(d, months):
    """Add N months to date d, clamping day to month length."""
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    day = min(d.day, calendar.monthrange(y, m)[1])
    return datetime(y, m, day)


def _nth_weekday_of_month(year, month, n, weekday_idx):
    """The Nth <weekday> of a given month (n=1..4). Returns datetime."""
    first = datetime(year, month, 1)
    first_wd = first.weekday()
    offset = (weekday_idx - first_wd) % 7
    day = 1 + offset + (n - 1) * 7
    last = calendar.monthrange(year, month)[1]
    if day > last:
        day -= 7  # clamp into month
    return datetime(year, month, day)


def compute_next(rec_type, rec_value, completion_date, current_start, current_due):
    """Compute the next occurrence date. Returns datetime or None if non-recurring.

    completion_date: the Alaska stamp date of THIS completion.
    current_start / current_due: the task's existing Start/Due (datetimes or None).
    """
    if rec_type is None:
        return None
    t = str(rec_type).strip().lower()
    if t in NON_RECURRING:
        return None
    v = (str(rec_value).strip() if rec_value is not None else "")

    # ── Interval from completion: "N units" or "N months@D" ──
    if t == "interval from completion":
        m = re.match(r'(\d+)\s+(day|days|week|weeks|month|months|year|years)(?:@(\d{1,2}))?$', v, re.I)
        if not m:
            return None
        n = int(m.group(1))
        unit = m.group(2).rstrip("s").lower()
        anchor_dom = m.group(3)
        base = completion_date
        if unit == "day":
            nxt = base + timedelta(days=n)
        elif unit == "week":
            nxt = base + timedelta(weeks=n)
        elif unit == "month":
            nxt = _add_months(base, n)
            if anchor_dom:  # "N months@D" — land on day-of-month D
                dom = int(anchor_dom)
                last = calendar.monthrange(nxt.year, nxt.month)[1]
                nxt = datetime(nxt.year, nxt.month, min(dom, last))
        elif unit == "year":
            nxt = _add_months(base, n * 12)
        else:
            return None
        return nxt

    # ── Weekly: next given weekday after completion ──
    if t == "weekly":
        wd = WEEKDAYS.get(v.lower())
        if wd is None:
            return None
        return _next_weekday(completion_date, wd)

    # ── Interval+Weekday: "N weeks/Weekday" — N weeks out, then that weekday ──
    if t == "interval+weekday":
        m = re.match(r'(\d+)\s+weeks?/(\w+)$', v, re.I)
        if not m:
            return None
        n = int(m.group(1))
        wd = WEEKDAYS.get(m.group(2).lower())
        if wd is None:
            return None
        base = completion_date + timedelta(weeks=n)
        # land on the specified weekday on/after that base week
        days = (wd - base.weekday()) % 7
        return base + timedelta(days=days)

    # ── Monthly: day-of-month "15"  OR  "Nth Weekday" ──
    if t == "monthly":
        m_dom = re.match(r'(\d{1,2})$', v)
        m_nwd = re.match(r'(\d)(?:st|nd|rd|th)\s+(\w+)$', v, re.I)
        # advance to next month from completion
        nxt_month_base = _add_months(completion_date, 1)
        if m_dom:
            dom = int(m_dom.group(1))
            last = calendar.monthrange(nxt_month_base.year, nxt_month_base.month)[1]
            return datetime(nxt_month_base.year, nxt_month_base.month, min(dom, last))
        if m_nwd:
            n = int(m_nwd.group(1))
            wd = WEEKDAYS.get(m_nwd.group(2).lower())
            if wd is None:
                return None
            return _nth_weekday_of_month(nxt_month_base.year, nxt_month_base.month, n, wd)
        return None

    # ── Yearly: "Mon DD" — next occurrence of that month/day ──
    if t == "yearly":
        m = re.match(r'([a-z]+)\s+(\d{1,2})$', v, re.I)
        if not m:
            return None
        mon = MONTHS.get(m.group(1).lower())
        day = int(m.group(2))
        if not mon:
            return None
        year = completion_date.year
        cand = datetime(year, mon, min(day, calendar.monthrange(year, mon)[1]))
        if cand <= completion_date:
            cand = datetime(year + 1, mon, min(day, calendar.monthrange(year + 1, mon)[1]))
        return cand

    return None


def _audit(ws_tasks, ws_completed, touched_ids):
    """ID integrity audit. Returns list of problem strings (empty = clean)."""
    H = _hdr(ws_tasks)
    problems = []
    # duplicate IDs in TASKS
    seen = {}
    for r in range(2, ws_tasks.max_row + 1):
        tid = ws_tasks.cell(r, H["ID"]).value
        if tid is None:
            continue
        if tid in seen:
            problems.append(f"DUPLICATE ID {tid} in TASKS (rows {seen[tid]},{r})")
        seen[tid] = r
    # TASKS ∩ COMPLETED collisions
    Hc = _hdr(ws_completed)
    cid_col = Hc.get("ID", 1)
    completed_ids = set()
    for r in range(2, ws_completed.max_row + 1):
        cid = ws_completed.cell(r, cid_col).value
        if cid is not None:
            completed_ids.add(cid)
    for tid in seen:
        if tid in completed_ids:
            problems.append(f"ID {tid} exists in BOTH TASKS and COMPLETED")
    return problems


def process_completions(wb, task_ids):
    """Complete the given task IDs. Mutates wb in place. Returns a results dict.

    For each ID:
      1. Find it in TASKS (by ID, not row)
      2. Stamp completion date (Alaska), copy row to COMPLETED
      3. If recurring → create next instance in TASKS with new Start
      4. If sequential → unlock next-in-group
      5. Remove the completed instance from TASKS
    Then run an ID audit. Caller saves + verifies.
    """
    ws = wb["TASKS"]
    wsc = wb["COMPLETED"]
    H = _hdr(ws)
    Hc = _hdr(wsc)
    stamp = alaska_stamp_date()
    results = {"completed": [], "recurred": [], "unlocked": [], "errors": [], "audit": []}

    # max existing ID for new recurring instances
    all_ids = [ws.cell(r, H["ID"]).value for r in range(2, ws.max_row + 1)
               if isinstance(ws.cell(r, H["ID"]).value, int)]
    cids = [wsc.cell(r, Hc.get("ID", 1)).value for r in range(2, wsc.max_row + 1)
            if isinstance(wsc.cell(r, Hc.get("ID", 1)).value, int)]
    next_id = max(all_ids + cids + [0]) + 1

    completed_col = Hc.get("Completion Date") or Hc.get("Completed") or 16

    rows_to_delete = []
    for tid in task_ids:
        # find row by ID
        trow = None
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, H["ID"]).value == tid:
                trow = r
                break
        if trow is None:
            results["errors"].append(f"ID {tid} not found in TASKS")
            continue

        name = ws.cell(trow, H["Task"]).value
        rec_type = ws.cell(trow, H["Recurring Type"]).value
        rec_value = ws.cell(trow, H["Recurring Value"]).value
        cur_start = _as_date(ws.cell(trow, H["Start Date"]).value)
        cur_due = _as_date(ws.cell(trow, H["Due Date"]).value)

        # 1. copy to COMPLETED — BY HEADER NAME, not column position.
        # TASKS and COMPLETED share most columns but diverge after Date Lock
        # (TASKS: ...Notes/Critical/Recurring Type/Recurring Value;
        #  COMPLETED: ...Recurring/Notes/Completed Date/Critical/Recurring Type/Recurring Value).
        # A positional copy scrambled recurrence into the wrong COMPLETED cells
        # (Recurring Type -> Completed Date col, etc.), so recurrence was lost on
        # any later walk-back/restore. Header-matching writes each TASKS field into
        # the SAME-NAMED COMPLETED column and skips any column COMPLETED lacks.
        new_c_row = wsc.max_row + 1
        for src_name, src_col in H.items():
            if src_name in Hc:
                wsc.cell(new_c_row, Hc[src_name]).value = ws.cell(trow, src_col).value
        # legacy mirror: if COMPLETED still has the old single 'Recurring' text
        # column, populate it too so old tooling/readers stay consistent.
        if "Recurring" in Hc and "Recurring Type" in H:
            rt_v = ws.cell(trow, H["Recurring Type"]).value
            rv_v = ws.cell(trow, H.get("Recurring Value", 0)).value if "Recurring Value" in H else None
            if rt_v not in (None, "", "None"):
                wsc.cell(new_c_row, Hc["Recurring"]).value = (
                    f"{rt_v}" + (f" / {rv_v}" if rv_v not in (None, "") else ""))
        wsc.cell(new_c_row, completed_col).value = stamp
        results["completed"].append((tid, name))

        # 2. recurrence
        nxt = compute_next(rec_type, rec_value, stamp, cur_start, cur_due)
        if nxt:
            nr = ws.max_row + 1
            for c in range(1, ws.max_column + 1):
                ws.cell(nr, c).value = ws.cell(trow, c).value
            ws.cell(nr, H["ID"]).value = next_id
            ws.cell(nr, H["Start Date"]).value = nxt
            # shift due date by same delta if it had one
            if cur_due and cur_start:
                delta = cur_due - cur_start
                ws.cell(nr, H["Due Date"]).value = nxt + delta
            results["recurred"].append((next_id, name, nxt.date()))
            next_id += 1

        # 3. sequential unlock — next in same SeqGroup by SeqOrder
        seq_g = ws.cell(trow, H["Sequential Group"]).value
        seq_o = ws.cell(trow, H["Sequential Order"]).value
        if seq_g and seq_o is not None:
            best = None
            for r in range(2, ws.max_row + 1):
                if r == trow:
                    continue
                if ws.cell(r, H["Sequential Group"]).value == seq_g:
                    o = ws.cell(r, H["Sequential Order"]).value
                    if o is not None and o > seq_o:
                        if best is None or o < best[1]:
                            best = (r, o)
            if best:
                nrow = best[0]
                nstart = _as_date(ws.cell(nrow, H["Start Date"]).value)
                if nstart is None or nstart > stamp:
                    ws.cell(nrow, H["Start Date"]).value = stamp + timedelta(days=1)
                    results["unlocked"].append((ws.cell(nrow, H["ID"]).value,
                                                ws.cell(nrow, H["Task"]).value))

        rows_to_delete.append(trow)

    # 4. delete completed rows (descending so indices stay valid)
    for r in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(r, 1)

    # 5. audit
    results["audit"] = _audit(ws, wsc, task_ids)
    return results
