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
# COURAGE stays manual: it routes a parent task to COMPLETED via the sacred
# process_completions pipeline, which does not live in the Action's world.
AUTO_SOURCES = {"ZONES", "MAINTENANCE", "SPIN_WHEEL"}


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
    # ID-BASED: uid = ZONES:{zid}. Find the row whose ZID column == zid.
    uid = comp.get("uid", "")
    zid = comp.get("zid")
    if zid is None and uid.startswith("ZONES:"):
        try:
            zid = int(uid.split(":")[1])
        except (IndexError, ValueError):
            zid = None
    if zid is None:
        return False, f"  ⚠️  malformed ZONE entry {uid}: no ZID"

    ws = wb["ZONES"]
    header = {c.value: i + 1 for i, c in enumerate(ws[1])}
    if "Completed" not in header or "ZID" not in header:
        return False, "  ❌ ZONES sheet missing Completed or ZID column"
    completed_col = header["Completed"]
    zid_col = header["ZID"]

    # Find the row by ZID (NOT by row number — rows shift, IDs don't)
    target_row = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, zid_col).value == zid:
            target_row = r
            break
    if target_row is None:
        return False, f"  ⚠️  ZID {zid} not found in ZONES — SKIP"

    actual_floor = ws.cell(target_row, header["Floor"]).value
    actual_zone = ws.cell(target_row, header["Zone"]).value

    existing = ws.cell(target_row, completed_col).value
    if existing is not None:
        return True, f"  ⏭️  {actual_floor}/{actual_zone} (ZID {zid}) already stamped — skip"

    ws.cell(target_row, completed_col).value = stamp
    return True, f"  ✅ {actual_floor}/{actual_zone} (ZID {zid}) stamped {stamp.date()}"


def process_maintenance_completion(wb, comp, stamp):
    """Stamp a MAINTENANCE sheet row.

    MAINTENANCE schema: col 1=Order, col 2=Task, col 3=Completed.
    uid format: MAINTENANCE:{row}
    """
    # ID-BASED: uid = MAINTENANCE:{mid}. Find row whose MID column == mid.
    uid = comp.get("uid", "")
    mid = comp.get("mid")
    if mid is None and uid.startswith("MAINTENANCE:"):
        try:
            mid = int(uid.split(":")[1])
        except (IndexError, ValueError):
            mid = None
    if mid is None:
        return False, f"  ⚠️  malformed MAINTENANCE entry {uid}: no MID"

    ws = wb["MAINTENANCE"]
    header = {c.value: i + 1 for i, c in enumerate(ws[1])}
    if "Completed" not in header or "MID" not in header:
        return False, "  ❌ MAINTENANCE sheet missing Completed or MID column"
    completed_col = header["Completed"]
    mid_col = header["MID"]
    task_col = header.get("Task", 2)

    target_row = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, mid_col).value == mid:
            target_row = r
            break
    if target_row is None:
        return False, f"  ⚠️  MID {mid} not found in MAINTENANCE — SKIP"

    actual_task = ws.cell(target_row, task_col).value
    existing = ws.cell(target_row, completed_col).value
    if existing is not None:
        return True, f"  ⏭️  MAINTENANCE '{actual_task}' (MID {mid}) already stamped — skip"

    ws.cell(target_row, completed_col).value = stamp
    return True, f"  ✅ MAINTENANCE '{actual_task}' (MID {mid}) stamped {stamp.date()}"


def process_spin_wheel_completions(wb, comps):
    """Delete SPIN WHEEL sheet rows for completed wheel items.

    uid format: SPIN_WHEEL:{row}. Deletes shift everything below up by one,
    so we MUST delete in descending row order or later deletes hit the wrong
    rows. Returns (processed_uids, messages).
    """
    if "SPIN WHEEL" not in wb.sheetnames:
        return [], ["  ❌ SPIN WHEEL sheet not found — skipping all spin completions"]

    ws = wb["SPIN WHEEL"]
    header = {c.value: i + 1 for i, c in enumerate(ws[1])}
    if "SID" not in header:
        return [], ["  ❌ SPIN WHEEL missing SID column — cannot ID-match"]
    sid_col = header["SID"]
    task_col = header.get("Task", 1)
    processed = []
    msgs = []

    # ID-BASED: uid = SPIN:{sid}. Resolve each SID to its CURRENT row, then
    # delete by descending row so shifts don't corrupt remaining deletes.
    # We resolve rows fresh (not from drain) so prior deletes in THIS batch
    # are already reflected by re-scanning each time.
    targets = []  # (sid, uid, label)
    for comp in comps:
        uid = comp.get("uid", "")
        sid = comp.get("sid")
        if sid is None and uid.startswith("SPIN:"):
            try:
                sid = int(uid.split(":")[1])
            except (IndexError, ValueError):
                sid = None
        if sid is None:
            msgs.append(f"  ⚠️  malformed SPIN entry {uid}: no SID")
            continue
        targets.append((sid, uid, comp.get("label", "?")))

    # Resolve all SIDs to rows FIRST (before any deletes), then delete descending.
    sid_to_row = {}
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, sid_col).value
        if v is not None:
            sid_to_row[v] = r

    resolved = []
    for sid, uid, label in targets:
        row = sid_to_row.get(sid)
        if row is None:
            msgs.append(f"  ⏭️  SID {sid} not found (already gone?) — clearing uid anyway")
            processed.append(uid)  # already absent = effectively done
            continue
        resolved.append((row, sid, uid, label))

    for row, sid, uid, label in sorted(resolved, key=lambda x: x[0], reverse=True):
        actual = ws.cell(row, task_col).value
        ws.delete_rows(row, 1)
        msgs.append(f"  ✅ SPIN deleted SID {sid} ('{actual or label}')")
        processed.append(uid)

    return processed, msgs


def write_sync_receipt(processed_uids, processed_details, remaining_completions, remaining_adds):
    """Write a token-free verification breadcrumb to the PUBLIC state repo.

    Chat-side Claude can read this via raw.githubusercontent.com with no token,
    so it can confirm 'did my tap file correctly?' without ever touching the
    private Beast. The Beast stays private; only this summary is public.

    Committed by the workflow yaml alongside drain.json.
    """
    receipt = {
        "lastRun": datetime.now(ZoneInfo("UTC")).isoformat(),
        "lastRunAlaska": datetime.now(ZoneInfo("America/Anchorage")).strftime("%Y-%m-%d %H:%M %Z"),
        "processedCount": len(processed_uids),
        "processed": processed_details,   # list of {uid, result} human-readable
        "remainingForClaude": {
            "completions": len(remaining_completions),
            "adds": len(remaining_adds),
        },
    }
    with open("last_sync.json", "w") as f:
        json.dump(receipt, f, indent=2)
    print(f"\n🧾 Sync receipt written: {len(processed_uids)} processed, "
          f"{len(remaining_completions)} completions + {len(remaining_adds)} adds left for Claude")


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
    pending_clear_from_chat = drain.get("pendingClear", [])
    print(f"📥 Drain has {len(completions)} completions + {len(adds)} adds + {len(pending_clear_from_chat)} pendingClear uids")

    if not completions and not adds and not pending_clear_from_chat:
        print("✨ Drain is empty — exiting cleanly")
        return

    # 2. Partition by auto-handleable vs manual
    auto_zones = [c for c in completions if c.get("source") == "ZONES"]
    auto_maintenance = [c for c in completions if c.get("source") == "MAINTENANCE"]
    auto_spin = [c for c in completions if c.get("source") == "SPIN_WHEEL"]
    other_completions = [c for c in completions if c.get("source") not in AUTO_SOURCES]

    print(f"🔧 Auto-processable: {len(auto_zones)} ZONES, {len(auto_maintenance)} MAINTENANCE, {len(auto_spin)} SPIN_WHEEL")
    print(f"⏸️  Leaving for Claude: {len(other_completions)} other completions, {len(adds)} adds")

    if pending_clear_from_chat:
        print(f"📞 Chat-side pendingClear: {len(pending_clear_from_chat)} uid(s) to clear from KV")

    # 3. Load beast ONLY if we have auto-processable completions
    wb = None
    beast_sha = None
    if auto_zones or auto_maintenance or auto_spin:
        print(f"\n📂 Loading beast from {BEAST_REPO}/{BEAST_FILE}...")
        beast_bytes, beast_sha = load_beast()
        print(f"   {len(beast_bytes)} bytes, SHA {beast_sha[:12]}")
        wb = load_workbook(io.BytesIO(beast_bytes))
    stamp = alaska_stamp_date() if wb else None
    if wb:
        print(f"   Stamp date (Alaska): {stamp.date()}")

    # 4a. Process ZONE completions
    processed_uids = []
    processed_details = []  # human-readable lines for the public receipt
    if auto_zones and wb:
        print(f"\n🏠 Processing {len(auto_zones)} ZONE completions:")
        for comp in auto_zones:
            ok, msg = process_zone_completion(wb, comp, stamp)
            print(msg)
            if ok:
                processed_uids.append(comp.get("uid"))
                processed_details.append({"uid": comp.get("uid"), "result": msg.strip()})

    # 4b. Process MAINTENANCE completions
    if auto_maintenance and wb:
        print(f"\n🔧 Processing {len(auto_maintenance)} MAINTENANCE completions:")
        for comp in auto_maintenance:
            ok, msg = process_maintenance_completion(wb, comp, stamp)
            print(msg)
            if ok:
                processed_uids.append(comp.get("uid"))
                processed_details.append({"uid": comp.get("uid"), "result": msg.strip()})

    # 4c. Process SPIN_WHEEL completions (row deletes, descending order)
    if auto_spin and wb:
        print(f"\n🎡 Processing {len(auto_spin)} SPIN_WHEEL completions:")
        spin_uids, spin_msgs = process_spin_wheel_completions(wb, auto_spin)
        for m in spin_msgs:
            print(m)
        processed_uids.extend(spin_uids)
        for u, m in zip(spin_uids, [x for x in spin_msgs if x.strip().startswith("✅")]):
            processed_details.append({"uid": u, "result": m.strip()})

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

    # 8. Write public sync receipt (token-free verification for chat-side Claude)
    write_sync_receipt(processed_uids, processed_details, new_completions, adds)


if __name__ == "__main__":
    main()
