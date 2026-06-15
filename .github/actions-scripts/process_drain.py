"""
process_drain.py — runs inside GitHub Action

Reads drain.json (this repo), processes each completion against beast.xlsx
(in karaemilie/hiveBrain), writes beast back, clears drain.json.

After successful processing, calls the worker's /clear-uids endpoint to
remove processed items from KV (F9: prevents ghost re-injection).

Also processes drain.json's `pendingClear` field if present — that's how
Claude (chat-side) signals "please clear these uids from KV." Chat writes
the uids to that field; Action does the network call (Claude container
can't reach workers.dev).

MVP scope: ZONES + MAINTENANCE auto-process. SPIN_WHEEL/TASKS/COURAGE
completions stay queued for Claude.
"""

import os
import sys
import json
import base64
import urllib.request
import urllib.error
import io
from datetime import datetime
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

# ─── CONFIG ──────────────────────────────────────────────────────
TOKEN = os.environ.get("CROSS_REPO_TOKEN", "")
if not TOKEN:
    print("❌ FATAL: CROSS_REPO_TOKEN env var not set")
    sys.exit(1)

BEAST_REPO = "karaemilie/hiveBrain"
BEAST_FILE = "masterHiveBrain.xlsx"
BEAST_BRANCH = "main"

DRAIN_PATH = "drain.json"  # in current repo (checkout)

WORKER_URL = "https://chaos-lotus-kv.theodidact.workers.dev"

# Sources this MVP handles automatically. All others stay queued for Claude.
AUTO_SOURCES = {"ZONES", "MAINTENANCE"}


# ─── ALASKA STAMP ────────────────────────────────────────────────
def alaska_stamp_date():
    """Return today's date in Alaska time (with 3am pivot)."""
    now_ak = datetime.now(ZoneInfo("America/Anchorage"))
    if now_ak.hour < 3:
        from datetime import timedelta
        return (now_ak - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)
    return now_ak.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)


# ─── GITHUB API ──────────────────────────────────────────────────
def gh(method, path, body=None):
    url = f"https://api.github.com{path}"
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "process-drain-action",
    }
    data = json.dumps(body).encode() if body else None
    if data:
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.status, json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read().decode())
        except Exception:
            body = {"error": "could not parse"}
        return e.code, body


def call_worker_clear_uids(uids):
    """Call worker POST /clear-uids to remove specific items from KV.

    Returns (status, response_dict). Best-effort — caller decides if a
    failure here is fatal (usually not — beast state is already correct,
    KV mismatch is recoverable).
    """
    if not uids:
        return 0, {"note": "no uids to clear"}
    body = json.dumps({"uids": list(uids)}).encode()
    req = urllib.request.Request(
        f"{WORKER_URL}/clear-uids",
        data=body,
        headers={
            "Content-Type": "application/json",
            "User-Agent": "process-drain-action",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.status, json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode())
        except Exception:
            return e.code, {"error": "could not parse"}
    except Exception as e:
        return 0, {"error": str(e)}


# ─── LOAD/SAVE BEAST ─────────────────────────────────────────────
def load_beast():
    status, meta = gh("GET", f"/repos/{BEAST_REPO}/contents/{BEAST_FILE}?ref={BEAST_BRANCH}")
    if status != 200:
        raise RuntimeError(f"load_beast failed: HTTP {status} — {meta}")
    return base64.b64decode(meta["content"]), meta["sha"]


def save_beast(beast_bytes, sha, commit_msg):
    body = {
        "message": commit_msg,
        "content": base64.b64encode(beast_bytes).decode(),
        "branch": BEAST_BRANCH,
        "sha": sha,
    }
    status, result = gh("PUT", f"/repos/{BEAST_REPO}/contents/{BEAST_FILE}", body)
    if status not in (200, 201):
        raise RuntimeError(f"save_beast failed: HTTP {status} — {result}")
    return result


# ─── ZONE PROCESSING ─────────────────────────────────────────────
def process_zone_completion(wb, comp, stamp):
    """Stamp a ZONES sheet cell. Returns (ok, message)."""
    floor = comp.get("floor")
    row = comp.get("row")
    label = comp.get("label", "?")
    if not floor or not row:
        return False, f"  ⚠️  malformed ZONE entry {comp.get('uid')}: missing floor/row"

    ws = wb["ZONES"]
    # Find the Completed column via header lookup (schema-safe)
    header = {c.value: i + 1 for i, c in enumerate(ws[1])}
    if "Completed" not in header:
        return False, "  ❌ ZONES sheet missing Completed column"
    completed_col = header["Completed"]

    # Verify row matches the floor and name
    actual_floor = ws.cell(row, 1).value
    actual_zone = ws.cell(row, 2).value
    if actual_floor != floor:
        return False, f"  ⚠️  row {row} is floor '{actual_floor}', expected '{floor}' — SKIP"

    # Idempotent check: already stamped?
    existing = ws.cell(row, completed_col).value
    if existing is not None:
        return True, f"  ⏭️  {floor}/{actual_zone} (row {row}) already stamped — skip"

    # Stamp it
    ws.cell(row, completed_col).value = stamp
    return True, f"  ✅ {floor}/{actual_zone} (row {row}) stamped {stamp.date()}"


def process_maintenance_completion(wb, comp, stamp):
    """Stamp a MAINTENANCE sheet row.

    MAINTENANCE schema: col 1=Order, col 2=Task, col 3=Completed.
    uid format: MAINTENANCE:{row}
    """
    row = comp.get("row")
    label = comp.get("label", "?")
    if not row:
        return False, f"  ⚠️  malformed MAINTENANCE entry {comp.get('uid')}: missing row"

    ws = wb["MAINTENANCE"]
    header = {c.value: i + 1 for i, c in enumerate(ws[1])}
    if "Completed" not in header:
        return False, "  ❌ MAINTENANCE sheet missing Completed column"
    completed_col = header["Completed"]

    actual_task = ws.cell(row, 2).value

    # Idempotent check
    existing = ws.cell(row, completed_col).value
    if existing is not None:
        return True, f"  ⏭️  MAINTENANCE '{actual_task}' (row {row}) already stamped — skip"

    ws.cell(row, completed_col).value = stamp
    return True, f"  ✅ MAINTENANCE '{actual_task}' (row {row}) stamped {stamp.date()}"


# ─── MAIN ────────────────────────────────────────────────────────
def main():
    # 1. Read drain.json from local checkout
    try:
        with open(DRAIN_PATH, "r") as f:
            drain = json.load(f)
    except FileNotFoundError:
        print(f"⚠️  {DRAIN_PATH} not found — nothing to process")
        return

    completions = drain.get("completions", [])
    adds = drain.get("adds", [])
    print(f"📥 Drain has {len(completions)} completions + {len(adds)} adds")

    if not completions and not adds:
        print("✨ Drain is empty — exiting cleanly")
        return

    # 2. Partition by auto-handleable vs manual
    auto_zones = [c for c in completions if c.get("source") == "ZONES"]
    auto_maintenance = [c for c in completions if c.get("source") == "MAINTENANCE"]
    other_completions = [c for c in completions if c.get("source") not in AUTO_SOURCES]

    print(f"🔧 Auto-processable: {len(auto_zones)} ZONES, {len(auto_maintenance)} MAINTENANCE")
    print(f"⏸️  Leaving for Claude: {len(other_completions)} other completions, {len(adds)} adds")

    pending_clear_from_chat = drain.get("pendingClear", [])
    if pending_clear_from_chat:
        print(f"📞 Chat-side pendingClear: {len(pending_clear_from_chat)} uid(s) to clear from KV")

    has_work = bool(auto_zones or auto_maintenance or pending_clear_from_chat)
    if not has_work:
        print("Nothing to do (no auto-completions, no pendingClear) — exiting")
        return

    # 3. Load beast ONLY if we have auto-processable completions
    wb = None
    beast_sha = None
    if auto_zones or auto_maintenance:
        print(f"\n📂 Loading beast from {BEAST_REPO}/{BEAST_FILE}...")
        beast_bytes, beast_sha = load_beast()
        print(f"   {len(beast_bytes)} bytes, SHA {beast_sha[:12]}")
        wb = load_workbook(io.BytesIO(beast_bytes))
    stamp = alaska_stamp_date() if wb else None
    if wb:
        print(f"   Stamp date (Alaska): {stamp.date()}")

    # 4a. Process ZONE completions
    processed_uids = []
    if auto_zones and wb:
        print(f"\n🏠 Processing {len(auto_zones)} ZONE completions:")
        for comp in auto_zones:
            ok, msg = process_zone_completion(wb, comp, stamp)
            print(msg)
            if ok:
                processed_uids.append(comp.get("uid"))

    # 4b. Process MAINTENANCE completions
    if auto_maintenance and wb:
        print(f"\n🔧 Processing {len(auto_maintenance)} MAINTENANCE completions:")
        for comp in auto_maintenance:
            ok, msg = process_maintenance_completion(wb, comp, stamp)
            print(msg)
            if ok:
                processed_uids.append(comp.get("uid"))

    if not processed_uids:
        print("\n⏭️  No auto-completions applied this run")
    else:
        # 5. Save beast
        print(f"\n💾 Saving beast back to github...")
        buf = io.BytesIO()
        wb.save(buf)
        new_beast_bytes = buf.getvalue()
        commit_msg = f"🤖 Auto-process: {len(processed_uids)} completion(s) stamped"
        result = save_beast(new_beast_bytes, beast_sha, commit_msg)
        print(f"   Committed: {result['commit']['sha'][:12]}")

    # 6. Determine what to clear from KV:
    #    - everything we just processed
    #    - everything chat-side flagged via drain.json's `pendingClear` field
    all_clear_uids = list(set(processed_uids) | set(pending_clear_from_chat))

    if all_clear_uids:
        print(f"\n🧹 Calling worker /clear-uids for {len(all_clear_uids)} uid(s):")
        for u in all_clear_uids:
            print(f"   - {u}")
        clear_status, clear_result = call_worker_clear_uids(all_clear_uids)
        print(f"   Status: {clear_status}")
        if clear_status == 200:
            print(f"   ✅ Removed: {clear_result.get('removed')} | Remaining: {clear_result.get('remaining')}")
        else:
            print(f"   ⚠️  Worker /clear-uids non-200: {clear_result}")
            print(f"   Beast state is still correct; KV may have ghosts until next run.")

    # 7. Rewrite drain.json — remove processed items, drop pendingClear field
    new_completions = [c for c in completions if c.get("uid") not in processed_uids]
    new_drain = {
        **drain,
        "completions": new_completions,
        "adds": adds,  # unchanged — still for Claude
        "processedAt": datetime.now(ZoneInfo("UTC")).isoformat(),
    }
    if processed_uids:
        new_drain["lastAutoProcessed"] = processed_uids
    # Remove pendingClear field once we've acted on it
    new_drain.pop("pendingClear", None)

    with open(DRAIN_PATH, "w") as f:
        json.dump(new_drain, f, indent=2)
    print(f"\n🧹 Drain rewritten: {len(new_completions)} completions remaining, {len(adds)} adds remaining")
    print(f"   (workflow yaml will commit drain.json change)")


if __name__ == "__main__":
    main()
