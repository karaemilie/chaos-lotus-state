"""
chaos_kv_refill.py — Autonomous wheel refill (runs inside GitHub Action)

When a completion is processed, the wheel should immediately get a fresh item
in that same category — finish a zone, the next zone appears; finish a Daily
Ten task, a new one drops in; same for Courage and Spin.

This module computes the replacement item(s) and returns them so the caller
(process_drain.py) can append them to state.json in the same run.

DESIGN PRINCIPLES
-----------------
1. ID-FIRST. Every emitted item carries a stable ID-based uid:
     TASKS:{id} · COURAGE:{parentId}:0 · ZONES:{zid} · MAINTENANCE:{mid} · SPIN:{sid}
   Never row-based. Rows shift; IDs don't.

2. NO DUPLICATES. Refill is given the set of uids already on the wheel and
   never re-adds one that's present.

3. SAME SELECTION RULES as the chat-side seed builder (chaos_kv_seed.py):
     • Daily Ten  → AW 3-7, not already on wheel, random (daily-stable seed)
     • Courage    → AW 1-2, most-overdue first, fallback oldest-start
     • Zones      → next uncompleted zone on the SAME floor (picker order)
     • Maintenance→ next uncompleted MAINTENANCE row
     • Spin       → next SPIN WHEEL item not already on the wheel
   These are kept deliberately simple and self-contained so the Action has no
   dependency on /mnt/project (which it cannot see).

4. EMOJI parity. Items carry an `emoji` field matching the seed builder so the
   wheel UI renders them identically.

The caller decides whether to refill 1:1 (one out, one in) or to a target
count. Default is 1:1 — replace exactly what was completed.
"""

import random
from datetime import date, datetime


# Floor emoji map — mirror of FLOOR_EMOJI.py (kept local; Action can't import it)
FLOOR_EMOJI = {
    "Upstairs": "🧺", "Main Floor": "🧺", "Basement": "🧺",
    "Digital": "💻", "Plant": "🌱", "Personal": "👱‍♀️",
    "Maintenance": "🏠", "Business": "💸", "Frog": "🐸",
}

DAILY_AW_RANGE = (3, 7)
COURAGE_AW_RANGE = (1, 2)

# Priority sort rank — lower = more urgent. Used for dry-pool fallback.
PRIORITY_RANK = {
    "ASAP": 0, "High": 1, "Soon": 2, "Normal": 3, "Medium": 3,
    "Sometime": 4, "Someday": 5, "Low": 6,
}


def _priority_rank(pri):
    return PRIORITY_RANK.get(str(pri).strip() if pri else "", 3)


def _future_tasks_in_band(wb, today, lo, hi, exclude_uids, uid_prefix):
    """Reservoir fallback: tasks with Start > today in the given AW band, not
    already excluded. Sorted by soonest start, then highest priority.

    uid_prefix is 'TASKS' or 'COURAGE' so we exclude the right uid shape.
    Returns sorted list of task dicts.
    """
    ws = wb["TASKS"]
    H = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    out = []
    for r in range(2, ws.max_row + 1):
        task = ws.cell(r, H["Task"]).value
        if not task:
            continue
        start = ws.cell(r, H["Start Date"]).value
        if hasattr(start, "date"):
            start = start.date()
        if start is None or start <= today:
            continue  # only FUTURE tasks here (available ones handled elsewhere)
        aw = ws.cell(r, H["Anchor Weight"]).value
        try:
            aw = int(aw)
        except (TypeError, ValueError):
            continue
        if not (lo <= aw <= hi):
            continue
        tid = ws.cell(r, H["ID"]).value
        if f"{uid_prefix}:{tid}" in exclude_uids or f"TASKS:{tid}" in exclude_uids or f"COURAGE:{tid}:0" in exclude_uids:
            continue
        due = ws.cell(r, H["Due Date"]).value
        if hasattr(due, "date"):
            due = due.date()
        out.append({
            "id": tid, "label": task, "aw": aw,
            "pri": ws.cell(r, H["Priority"]).value,
            "dur": ws.cell(r, H["Duration (min)"]).value,
            "ml": ws.cell(r, H["Mental Load"]).value,
            "proj": ws.cell(r, H["Project"]).value,
            "cat": ws.cell(r, H["Category"]).value,
            "start": start, "due": due,
            "critical": bool(ws.cell(r, H["Critical"]).value),
        })
    # soonest start first, then highest priority (lowest rank)
    out.sort(key=lambda t: (t["start"], _priority_rank(t["pri"])))
    return out


# ─── HEADER HELPERS ──────────────────────────────────────────────
def _hdr(ws):
    return {c.value: i + 1 for i, c in enumerate(ws[1])}


def _as_date(v):
    if hasattr(v, "date"):
        return v.date()
    return v


# ─── ZONE REFILL ─────────────────────────────────────────────────
def refill_zone(wb, completed_comp, on_wheel_uids):
    """After a zone on `floor` is completed, return the next uncompleted zone
    on that SAME floor (by ZONES sheet order), as a wheel item dict — or None.

    `completed_comp` is the drain completion that was just processed; we read
    its floor from the Beast via the just-stamped ZID (most reliable).
    """
    ws = wb["ZONES"]
    H = _hdr(ws)
    if "ZID" not in H:
        return None
    zid_col, floor_col, zone_col, comp_col = H["ZID"], H["Floor"], H["Zone"], H["Completed"]

    # Resolve the completed zone's floor from its ZID
    comp_zid = completed_comp.get("zid")
    uid = completed_comp.get("uid", "")
    if comp_zid is None and uid.startswith("ZONES:"):
        try:
            comp_zid = int(uid.split(":")[1])
        except (IndexError, ValueError):
            comp_zid = None
    if comp_zid is None:
        return None

    floor = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, zid_col).value == comp_zid:
            floor = ws.cell(r, floor_col).value
            break
    if floor is None:
        return None

    # Walk that floor top-to-bottom for the first uncompleted zone not already on the wheel
    emoji = FLOOR_EMOJI.get(floor, "📍")
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, floor_col).value != floor:
            continue
        if ws.cell(r, comp_col).value is not None:
            continue
        zid = ws.cell(r, zid_col).value
        if zid is None:
            continue
        new_uid = f"ZONES:{zid}"
        if new_uid in on_wheel_uids:
            continue
        zone_name = ws.cell(r, zone_col).value
        return {
            "source": "ZONES", "zid": zid, "floor": floor,
            "label": f"{emoji} {zone_name}", "emoji": emoji,
            "zoneName": zone_name, "uid": new_uid,
        }

    # FLOOR EXHAUSTED: every zone on this floor is completed. Wrap to the top
    # (re-serve the first zone) so the wheel stays full, AND signal that this
    # floor needs a proper reset (clearing Completed dates) — which is picker
    # logic that must run Claude-side. We tag the returned item with
    # `_resetFloor` so the caller can flag it in the drain queue.
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, floor_col).value != floor:
            continue
        zid = ws.cell(r, zid_col).value
        if zid is None:
            continue
        new_uid = f"ZONES:{zid}"
        if new_uid in on_wheel_uids:
            continue  # already showing — skip to keep it distinct
        zone_name = ws.cell(r, zone_col).value
        return {
            "source": "ZONES", "zid": zid, "floor": floor,
            "label": f"{emoji} {zone_name}", "emoji": emoji,
            "zoneName": zone_name, "uid": new_uid,
            "_resetFloor": floor,  # signal: this floor needs a Claude-side reset
        }
    return None  # truly nothing left (every zone already on wheel)


# ─── MAINTENANCE REFILL ──────────────────────────────────────────
def refill_maintenance(wb, on_wheel_uids):
    """Return the next uncompleted MAINTENANCE row not already on the wheel."""
    ws = wb["MAINTENANCE"]
    H = _hdr(ws)
    if "MID" not in H:
        return None
    mid_col, task_col, comp_col = H["MID"], H.get("Task", 2), H["Completed"]
    emoji = FLOOR_EMOJI["Maintenance"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, comp_col).value is not None:
            continue
        mid = ws.cell(r, mid_col).value
        task = ws.cell(r, task_col).value
        if mid is None or not task:
            continue
        new_uid = f"MAINTENANCE:{mid}"
        if new_uid in on_wheel_uids:
            continue
        return {
            "source": "MAINTENANCE", "mid": mid,
            "label": f"{emoji} {task}", "emoji": emoji,
            "taskName": task, "uid": new_uid,
        }
    # ALL MAINTENANCE DONE: wrap to top + flag for Claude-side reset
    for r in range(2, ws.max_row + 1):
        mid = ws.cell(r, mid_col).value
        task = ws.cell(r, task_col).value
        if mid is None or not task:
            continue
        new_uid = f"MAINTENANCE:{mid}"
        if new_uid in on_wheel_uids:
            continue
        return {
            "source": "MAINTENANCE", "mid": mid,
            "label": f"{emoji} {task}", "emoji": emoji,
            "taskName": task, "uid": new_uid,
            "_resetFloor": "Maintenance",
        }
    return None


# ─── SPIN REFILL ─────────────────────────────────────────────────
def refill_spin(wb, on_wheel_uids):
    """Return the next SPIN WHEEL item (by SID) not already on the wheel."""
    ws = wb["SPIN WHEEL"]
    H = _hdr(ws)
    if "SID" not in H:
        return None
    sid_col, task_col = H["SID"], H.get("Task", 1)
    for r in range(2, ws.max_row + 1):
        sid = ws.cell(r, sid_col).value
        task = ws.cell(r, task_col).value
        if sid is None or not task:
            continue
        new_uid = f"SPIN:{sid}"
        if new_uid in on_wheel_uids:
            continue
        return {
            "source": "SPIN_WHEEL", "sid": sid,
            "label": task, "uid": new_uid,
        }
    return None  # all spin items already on wheel


# ─── TASKS (DAILY TEN) REFILL ────────────────────────────────────
def _available_tasks(wb, today):
    """All TASKS rows with Start <= today and a numeric AW. Returns list of dicts."""
    ws = wb["TASKS"]
    H = _hdr(ws)
    out = []
    for r in range(2, ws.max_row + 1):
        task = ws.cell(r, H["Task"]).value
        if not task:
            continue
        start = _as_date(ws.cell(r, H["Start Date"]).value)
        if start is None or start > today:
            continue
        aw = ws.cell(r, H["Anchor Weight"]).value
        try:
            aw = int(aw)
        except (TypeError, ValueError):
            continue
        due = _as_date(ws.cell(r, H["Due Date"]).value)
        out.append({
            "id": ws.cell(r, H["ID"]).value, "label": task, "aw": aw,
            "pri": ws.cell(r, H["Priority"]).value,
            "dur": ws.cell(r, H["Duration (min)"]).value,
            "ml": ws.cell(r, H["Mental Load"]).value,
            "proj": ws.cell(r, H["Project"]).value,
            "cat": ws.cell(r, H["Category"]).value,
            "notes": ws.cell(r, H["Notes"]).value,
            "start": start, "due": due,
            "critical": bool(ws.cell(r, H["Critical"]).value),
        })
    return out


def refill_daily_ten(wb, on_wheel_uids, today):
    """Pick one new AW 3-7 task not already on the wheel. Daily-stable random.
    If the available pool is empty, fall back to the future reservoir
    (soonest start + highest priority) so the wheel never runs dry."""
    pool = [t for t in _available_tasks(wb, today)
            if DAILY_AW_RANGE[0] <= t["aw"] <= DAILY_AW_RANGE[1]
            and f"TASKS:{t['id']}" not in on_wheel_uids]
    if pool:
        rng = random.Random(today.toordinal() + len(on_wheel_uids))
        t = rng.choice(pool)
    else:
        # DRY FALLBACK: reach into future-dated tasks, soonest+highest-priority
        future = _future_tasks_in_band(wb, today, *DAILY_AW_RANGE, on_wheel_uids, "TASKS")
        if not future:
            return None
        t = future[0]
    return {
        "source": "TASKS", "id": t["id"], "label": t["label"], "aw": t["aw"],
        "pri": t["pri"], "dur": t["dur"], "ml": t["ml"], "proj": t["proj"],
        "cat": t["cat"], "critical": t["critical"], "uid": f"TASKS:{t['id']}",
    }


# ─── COURAGE REFILL ──────────────────────────────────────────────
def refill_courage(wb, on_wheel_uids, today):
    """Pick one new AW 1-2 task (most overdue first) not already on the wheel.

    NEW MODEL: no pre-drafted micro. The 🔥 flame is the user's cue to brainstorm
    a first step themselves, in the moment, on the wheel. Tapping it done =
    "I faced it and cracked it open." The parent completes via the engine; any
    first step the user brainstorms becomes a SPIN one-off if they choose.

    So a Courage item is just the task itself, labeled with 🔥 as the
    decompose-cue. No micro text, no _needsMicro flag.
    """
    pool = [t for t in _available_tasks(wb, today)
            if COURAGE_AW_RANGE[0] <= t["aw"] <= COURAGE_AW_RANGE[1]
            and f"COURAGE:{t['id']}:0" not in on_wheel_uids
            and f"TASKS:{t['id']}" not in on_wheel_uids]
    if pool:
        overdue = [t for t in pool if t["due"] and t["due"] < today]
        overdue.sort(key=lambda t: (today - t["due"]).days, reverse=True)
        pick = overdue[0] if overdue else sorted(pool, key=lambda t: t["start"])[0]
    else:
        # DRY FALLBACK: future AW1-2 tasks, soonest start + highest priority
        future = _future_tasks_in_band(wb, today, *COURAGE_AW_RANGE, on_wheel_uids, "COURAGE")
        if not future:
            return None
        pick = future[0]
    return {
        "source": "COURAGE", "parentId": pick["id"], "stepIndex": 0,
        "parentLabel": pick["label"], "aw": pick["aw"],
        "due": pick["due"].isoformat() if pick["due"] else None,
        "label": f"🔥 {pick['label']}",  # flame = decompose-cue, not a micro
        "emoji": "🔥", "uid": f"COURAGE:{pick['id']}:0",
    }


# ─── DISPATCH ────────────────────────────────────────────────────
def refill_for(source, wb, completed_comp, on_wheel_uids, today):
    """Route a completed item's source to the right refill function.

    Returns a single new item dict (or None if nothing eligible).
    """
    if source == "ZONES":
        return refill_zone(wb, completed_comp, on_wheel_uids)
    if source == "MAINTENANCE":
        return refill_maintenance(wb, on_wheel_uids)
    if source == "SPIN_WHEEL":
        # Spin items are the full one-off pool, all shown at once. Completing
        # one just removes it — the pool shrinks until new ones are added.
        # No refill by design.
        return None
    if source == "TASKS":
        return refill_daily_ten(wb, on_wheel_uids, today)
    if source == "COURAGE":
        return refill_courage(wb, on_wheel_uids, today)
    return None
