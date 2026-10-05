#!/usr/bin/env python3
"""
zendesk_flexera.py
------------------
ZENDESK API -> FLEXERA (Snow Atlas) SaaS DATA IMPORT, in one run.

  1. Pull every agent / admin from the Zendesk Users API (full roster)
  2. Build the two Flexera CSVs in memory-safe text form (IDs never
     touch Excel, so no scientific-notation problem)
  3. Push each file: token -> request upload URL -> PUT to blob -> import

Flexera mappings this script writes for
  user assignments (8fc88134...)          user activities (e5f01062...)
    Email address        = email            Unique user identifier = id
    Unique user id       = id               Activity or event name = last_login_at
    Subscription name    = custom_role_name Activity recorded value = last_login_at
    Display name         = name
    Username             = name
    Account creation     = created_at
    Account last updated = last_login_at

Users who have never logged in, or who have no email (e.g. the "AI Agent"
bot), are still sent. Flexera may warn on those rows and skip only them.
Suspended users are sent as-is.

Conditions come from ZENDESK_DEFAULT_ARGS in .env (same line the extract
script uses), or from the command line, which replaces them completely:

    --all-time                     full current roster (default)
    --start / --end YYYY-MM-DD     only users whose --date-field is in the window
    --last-days N                  trailing window
    --date-field created|updated   which timestamp the window uses (default: created)
    --exclude-seat-class "A,B"     leave out roles; matched case-insensitively
                                   against custom_role_name OR seat_class
    --include-deleted              also send users with active == False
    --dry-run                      build the CSVs, send nothing

    --check-ticket-activity        also look up each user's ticket status in Zendesk
    --inactivity-days N            ticket look-back window in days (default 90)

Ticket status (tickets_assigned, tickets_submitted, last_ticket_activity_at,
days_since_ticket_activity, no_tickets_90d) is written to
zendesk_users_snapshot.csv ONLY. It is never added to the Flexera files and
nothing about tickets is pushed to Flexera. It costs 2 Zendesk search calls per
user (4 for users with nothing in the window), so it adds a few minutes on a
large roster. If it fails part-way the Flexera push still goes ahead.

Flags that only mean something to the extract script (--exclude-end-users,
--inactive-only, --out, --active-only) are accepted and ignored, so one
ZENDESK_DEFAULT_ARGS line can serve both scripts.

Every run is logged to output/zendesk_flexera.log (see the LOGGING section).

WARNING: a time window sends Flexera only part of the roster. Licensed users
outside the window will be missing from that import.

.env (next to this script):
    ZENDESK_DOMAIN=yoursubdomain
    ZENDESK_EMAIL=you@company.com
    ZENDESK_API_TOKEN=...
    FLEXERA_CLIENT_ID=...
    FLEXERA_CLIENT_SECRET=...
    ZENDESK_DEFAULT_ARGS=--all-time --exclude-seat-class "Agente Light" --check-ticket-activity
    LOG_FILE=...                                (optional, default output/zendesk_flexera.log)

Usage:
    python zendesk_flexera.py                       # uses ZENDESK_DEFAULT_ARGS
    python zendesk_flexera.py --dry-run --all-time  # manual args replace the defaults
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import platform
import re
import shlex
import sys
import time
from datetime import datetime, timedelta, timezone
from logging.handlers import RotatingFileHandler

import requests
from dotenv import find_dotenv, load_dotenv

# --------------------------- CONFIGURATION ---------------------------

FLEXERA_BASE = "https://snowatlas-westeurope.flexera.eu"
TOKEN_URL    = f"{FLEXERA_BASE}/idp/api/connect/token"
UPLOAD_URL   = f"{FLEXERA_BASE}/api/saas/import/v1/uploads"

ASSIGNMENTS_CONFIG_ID = "8fc88134-1797-4f0f-bdfc-1e83ad00d384"   # user assignments
ACTIVITIES_CONFIG_ID  = "e5f01062-75ac-43fe-94cc-b79ac0b250ba"   # user activities

OUTPUT_DIR      = os.path.dirname(os.path.abspath(__file__))
SNAPSHOT_CSV    = os.path.join(OUTPUT_DIR, "zendesk_users_snapshot.csv")  # for review only
ASSIGNMENTS_CSV = os.path.join(OUTPUT_DIR, "zendesk_license.csv")
ACTIVITIES_CSV  = os.path.join(OUTPUT_DIR, "zendesk_activities.csv")

# Output headers - must match the CSV fields in each Flexera mapping EXACTLY
ASSIGNMENTS_TEMPLATE = ["email", "id", "custom_role_name", "name",
                        "created_at", "last_login_at"]
ACTIVITIES_TEMPLATE  = ["id", "last_login_at"]
SNAPSHOT_TEMPLATE    = ["id", "name", "email", "role", "role_type", "seat_class",
                        "custom_role_name", "active", "suspended",
                        "created_at", "updated_at", "last_login_at"]

ONLY_ACTIVE_USERS                 = True   # drop deleted users (active == False)
DROP_BLANK_ACTIVITY_ROWS          = False  # False = send never-logged-in users too (Flexera warns)
FILL_SUBSCRIPTION_FROM_SEAT_CLASS = True   # blank custom_role_name -> seat_class

API_TIMEOUT = 60
MAX_RETRIES = 5

SEARCH_PAUSE = 0.7        # ticket search is rate limited harder than the users endpoint
TICKET_COLUMNS = ["tickets_assigned", "tickets_submitted", "last_ticket_activity_at",
                  "days_since_ticket_activity", "no_tickets_90d"]
INACTIVITY_DAYS = 90      # default ticket window for --check-ticket-activity

ROLE_TYPE_LABELS = {
    0: "Custom agent", 1: "Light agent", 2: "Chat agent",
    3: "Contributor", 4: "Admin", 5: "Billing admin",
}

# ------------------------------ LOGGING ------------------------------
# Every run is appended to output/<script>.log (override with LOG_FILE in .env).
# Each run starts with a separator line and every line is timestamped. Whatever
# the script prints is also written to the log, except the per-user progress
# lines. Secrets are never logged: the API token/secret values and any URL query
# string (Flexera's upload URL carries a signed token) are redacted. When the
# file passes LOG_MAX_BYTES it is renamed to <name>.log.1 and a new one starts.

LOG_MAX_BYTES = 5 * 1024 * 1024
LOG_LABEL     = "zendesk_flexera"
_SECRET_ENV   = ("ZENDESK_API_TOKEN", "FLEXERA_CLIENT_SECRET")
_URL_QUERY    = re.compile(r"(https?://[^\s?'\"<>]+)\?[^\s'\"<>]*")
_PROGRESS     = re.compile(r"^\s*\[\d+/\d+\]")      # "  [3/26] name ..." live counter
log = logging.getLogger(LOG_LABEL)

def _redact(text: str) -> str:
    text = _URL_QUERY.sub(r"\1?<redacted>", text)
    for name in _SECRET_ENV:
        value = os.getenv(name, "")
        if len(value) >= 6:
            text = text.replace(value, "<redacted>")
    return text

class _RedactingFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        return _redact(super().format(record))

class _Tee:
    """Pass output straight through to the console and log each finished line."""
    def __init__(self, stream, level: int):
        self._stream, self._level, self._buf = stream, level, ""
    def write(self, text: str) -> int:
        self._stream.write(text)
        self._buf += text
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            line = line.rstrip()
            if line.strip() and not _PROGRESS.match(line):
                log.log(self._level, line)
        return len(text)
    def flush(self) -> None:
        self._stream.flush()
    def __getattr__(self, name):
        return getattr(self._stream, name)

def run_logged(main_fn, tee: dict) -> None:
    """Run main_fn with logging. tee maps 'stdout'/'stderr' to the level used for it."""
    path = os.getenv("LOG_FILE", "").strip() or os.path.join(OUTPUT_DIR, "output", f"{LOG_LABEL}.log")
    try:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        handler = RotatingFileHandler(path, maxBytes=LOG_MAX_BYTES, backupCount=1, encoding="utf-8")
    except OSError as exc:
        print(f"WARNING: cannot write log file {path}: {exc}. Continuing without a log.")
        main_fn()
        return
    handler.setFormatter(_RedactingFormatter("%(asctime)s %(levelname)-7s %(message)s",
                                             "%Y-%m-%d %H:%M:%S"))
    log.setLevel(logging.INFO)
    log.propagate = False
    log.addHandler(handler)
    print(f"Log file: {path}")

    real = {"stdout": sys.stdout, "stderr": sys.stderr}
    log.info("=" * 72)
    log.info("RUN START  %s  (python %s)", LOG_LABEL, platform.python_version())
    started = time.time()
    for name, level in tee.items():
        setattr(sys, name, _Tee(real[name], level))
    try:
        main_fn()
        log.info("RUN END    ok (%.1fs)", time.time() - started)
    except SystemExit as exc:
        if exc.code in (None, 0):
            log.info("RUN END    ok (%.1fs)", time.time() - started)
        else:
            reason = exc.code if isinstance(exc.code, str) else f"exit code {exc.code}"
            log.error("STOPPED   %s", reason)
        raise
    except KeyboardInterrupt:
        log.warning("RUN END    interrupted by user")
        raise
    except Exception:
        log.exception("CRASHED")
        raise
    finally:
        for name, stream in real.items():
            setattr(sys, name, stream)
        log.removeHandler(handler)
        handler.close()


# ------------------------------ SETTINGS -----------------------------

def sanitize_domain(raw: str) -> str:
    value = raw.strip()
    if "://" in value:
        value = value.split("://", 1)[1]
    value = value.strip("/").split("/")[0]
    if value.endswith(".zendesk.com"):
        value = value[:-len(".zendesk.com")]
    return value.strip()

def load_settings() -> dict[str, str]:
    dotenv_path = find_dotenv(usecwd=True) or os.path.join(OUTPUT_DIR, ".env")
    names = ["ZENDESK_DOMAIN", "ZENDESK_EMAIL", "ZENDESK_API_TOKEN",
             "FLEXERA_CLIENT_ID", "FLEXERA_CLIENT_SECRET"]
    cfg = {n: os.getenv(n, "").strip() for n in names}
    missing = [n for n, v in cfg.items() if not v]
    if missing:
        sys.exit(f"Missing in .env ({dotenv_path}): {', '.join(missing)}")
    # Older .env files may still use this; merged with --exclude-seat-class
    cfg["ZENDESK_EXCLUDE_ROLES"] = os.getenv("ZENDESK_EXCLUDE_ROLES", "").strip().strip('"').strip("'")
    cfg["ZENDESK_DOMAIN"] = sanitize_domain(cfg["ZENDESK_DOMAIN"])
    if not cfg["ZENDESK_DOMAIN"] or "." in cfg["ZENDESK_DOMAIN"]:
        sys.exit("ZENDESK_DOMAIN is invalid. Use 'yourcompany' or "
                 "'https://yourcompany.zendesk.com'.")
    return cfg


# ------------------------------ ZENDESK ------------------------------

def zd_get(session: requests.Session, url: str, params: dict | None = None) -> dict:
    """GET with 429 / 5xx retry, honouring Retry-After."""
    for attempt in range(1, MAX_RETRIES + 1):
        resp = session.get(url, params=params, timeout=API_TIMEOUT)
        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 60))
            print(f"  Zendesk rate limit, sleeping {wait}s")
            time.sleep(wait)
            continue
        if resp.status_code in (500, 502, 503, 504) and attempt < MAX_RETRIES:
            time.sleep(2 ** attempt)
            continue
        if resp.status_code == 401:
            sys.exit("Zendesk 401 Unauthorised. Check ZENDESK_EMAIL / ZENDESK_API_TOKEN "
                     "and that token access is enabled in Admin Center.")
        if resp.status_code == 403:
            sys.exit("Zendesk 403 Forbidden. The account needs admin rights.")
        resp.raise_for_status()
        return resp.json()
    sys.exit(f"Zendesk: gave up after {MAX_RETRIES} attempts on {url}")

def fetch_custom_roles(session: requests.Session, base: str) -> dict[int, str]:
    try:
        data = zd_get(session, f"{base}/api/v2/custom_roles")
    except requests.HTTPError:
        print("  custom roles unavailable on this plan, continuing")
        return {}
    return {r["id"]: r.get("name", "") for r in data.get("custom_roles", [])}

def fetch_agents(session: requests.Session, base: str) -> list[dict]:
    """Every agent and admin (end users filtered out by Zendesk)."""
    url: str | None = f"{base}/api/v2/users"
    params: dict | None = {"page[size]": 100, "role[]": ["agent", "admin"]}
    users: list[dict] = []
    page = 0
    while url:
        page += 1
        data = zd_get(session, url, params)
        batch = data.get("users", [])
        users.extend(batch)
        print(f"  page {page}: +{len(batch)} (total {len(users)})")
        if not (data.get("meta") or {}).get("has_more"):
            break
        url = (data.get("links") or {}).get("next")
        params = None
        time.sleep(0.3)
    return users

def seat_class(user: dict, custom_roles: dict[int, str]) -> tuple[str, str]:
    """Return (seat_class, custom_role_name)."""
    role = user.get("role") or ""
    role_type = user.get("role_type")
    crid = user.get("custom_role_id")
    custom_name = custom_roles.get(crid, "") if crid else ""
    if role == "end-user":
        return "End user", ""
    if role_type in ROLE_TYPE_LABELS:
        label = ROLE_TYPE_LABELS[role_type]
        if role_type == 0 and custom_name:
            label = f"Custom agent: {custom_name}"
        return label, custom_name
    return ("Admin" if role == "admin" else "Agent" if role == "agent" else role or "Unknown"), custom_name

def to_row(user: dict, custom_roles: dict[int, str]) -> dict[str, str]:
    cls, custom_name = seat_class(user, custom_roles)
    return {
        "id": str(user.get("id") or ""),          # text, exact digits
        "name": (user.get("name") or "").strip(),
        "email": (user.get("email") or "").strip().lower(),
        "role": user.get("role") or "",
        "role_type": "" if user.get("role_type") is None else str(user["role_type"]),
        "seat_class": cls,
        "custom_role_name": custom_name,
        "active": str(user.get("active")).upper(),
        "suspended": str(user.get("suspended")).upper(),
        "created_at": user.get("created_at") or "",
        "updated_at": user.get("updated_at") or "",
        "last_login_at": user.get("last_login_at") or "",
    }


def search_latest(session: requests.Session, base: str, query: str) -> tuple[int, datetime | None]:
    """One /search call returning (total_count, updated_at of the newest match)."""
    data = zd_get(session, f"{base}/api/v2/search.json", {
        "query": query, "sort_by": "updated_at", "sort_order": "desc", "per_page": 1,
    })
    results = data.get("results") or []
    latest = parse_ts(results[0].get("updated_at")) if results else None
    return int(data.get("count", 0)), latest

def _newest(*stamps: datetime | None) -> datetime | None:
    present = [t for t in stamps if t is not None]
    return max(present) if present else None

def ticket_activity(session: requests.Session, base: str, user_id: str,
                    since: datetime) -> tuple[int, int, datetime | None]:
    """(tickets_assigned, tickets_submitted, last_ticket_activity).

    Counts cover tickets touched since `since`. The last-activity date has no
    date limit; the unbounded lookup only runs when the window found nothing.
    Search covers assignee and submitter only - a Light agent who just adds
    comments will show zero here even if they work daily.
    """
    day = since.strftime("%Y-%m-%d")
    assigned, last_a = search_latest(session, base, f"type:ticket assignee:{user_id} updated>{day}")
    time.sleep(SEARCH_PAUSE)
    submitted, last_s = search_latest(session, base, f"type:ticket submitter:{user_id} updated>{day}")
    time.sleep(SEARCH_PAUSE)
    last = _newest(last_a, last_s)
    if last is None:
        _, last_a = search_latest(session, base, f"type:ticket assignee:{user_id}")
        time.sleep(SEARCH_PAUSE)
        _, last_s = search_latest(session, base, f"type:ticket submitter:{user_id}")
        time.sleep(SEARCH_PAUSE)
        last = _newest(last_a, last_s)
    return assigned, submitted, last

def add_ticket_status(session: requests.Session, base: str, rows: list[dict],
                      days: int, now: datetime) -> None:
    """Fill the ticket columns on each row. Never raises: a failure leaves blanks."""
    cutoff = now - timedelta(days=days)
    for r in rows:
        for col in TICKET_COLUMNS:
            r[col] = ""
    print(f"\nChecking ticket activity for {len(rows)} user(s) since {cutoff:%Y-%m-%d} "
          f"(at least {len(rows) * 2} Zendesk search calls)...")
    failures = consecutive = checked = 0
    for i, r in enumerate(rows, 1):
        try:
            assigned, submitted, last = ticket_activity(session, base, r["id"], cutoff)
        except (requests.RequestException, SystemExit) as exc:
            failures += 1
            consecutive += 1
            reason = exc.code if isinstance(exc, SystemExit) else exc
            log.warning("ticket lookup failed for user id %s: %s", r["id"], reason)
            print(f"  [{i}/{len(rows)}] {r['name'][:34]:<34} lookup failed")
            if consecutive >= 3:
                print("  3 lookups failed in a row - skipping ticket status for the rest.")
                break
            continue
        consecutive = 0
        checked += 1
        none_now = (assigned + submitted) == 0
        r["tickets_assigned"], r["tickets_submitted"] = assigned, submitted
        if last is not None:
            r["last_ticket_activity_at"] = last.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            r["days_since_ticket_activity"] = (now - last).days
        r["no_tickets_90d"] = "TRUE" if none_now else "FALSE"
        if not none_now:
            flag = f"{assigned + submitted} ticket(s)"
        elif last is not None:
            flag = f"no activity, last ticket {r['days_since_ticket_activity']}d ago"
        else:
            flag = "no tickets ever"
        print(f"  [{i}/{len(rows)}] {r['name'][:34]:<34} {flag}")

    no_tix = sum(1 for r in rows if r["no_tickets_90d"] == "TRUE")
    never = sum(1 for r in rows if r["no_tickets_90d"] == "TRUE" and r["days_since_ticket_activity"] == "")
    print(f"Ticket status: {checked} of {len(rows)} checked, {failures} failed.")
    print(f"  no ticket activity in {days} days: {no_tix}  ({never} have no tickets at all)")
    print("  Note: search covers assignee and submitter only; Light agents who just "
          "comment will show zero.")


# ------------------------------ FLEXERA ------------------------------

def get_access_token(client_id: str, client_secret: str) -> str:
    resp = requests.post(
        TOKEN_URL,
        data={"grant_type": "client_credentials",
              "client_id": client_id, "client_secret": client_secret},
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    resp.raise_for_status()
    print("Obtained Flexera access token.")
    return resp.json()["access_token"]

def request_upload_url(token: str) -> tuple[str, str]:
    resp = requests.post(UPLOAD_URL, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    print(f"  upload URL received (fileId={data.get('fileId')}, expires in {data.get('expires')}s)")
    return data["url"], data["fileId"]

def upload_csv_to_blob(presigned_url: str, csv_path: str) -> None:
    with open(csv_path, "rb") as f:
        resp = requests.put(presigned_url, data=f,
                            headers={"x-ms-blob-type": "BlockBlob", "Content-Type": "text/csv"},
                            timeout=120)
    resp.raise_for_status()
    print(f"  uploaded to blob storage (HTTP {resp.status_code})")

def trigger_import(token: str, file_id: str, config_id: str) -> None:
    url = f"{FLEXERA_BASE}/api/saas/import/v1/configs/{config_id}/records/import"
    resp = requests.post(url,
                         headers={"Authorization": f"Bearer {token}",
                                  "Content-Type": "application/json"},
                         data=json.dumps({"fileId": file_id}), timeout=60)
    resp.raise_for_status()
    print(f"  import task created (HTTP {resp.status_code}): {resp.text[:500]}")

def push_to_flexera(token: str, csv_path: str, config_id: str, label: str) -> None:
    print(f"\n--- Pushing {label} ({os.path.basename(csv_path)}) ---")
    url, file_id = request_upload_url(token)
    upload_csv_to_blob(url, csv_path)
    trigger_import(token, file_id, config_id)


# ------------------------------ PIPELINE -----------------------------

def write_csv(path: str, rows: list[dict], columns: list[str], bom: bool = False) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig" if bom else "utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)

EXTRACT_ONLY_FLAGS = {"--exclude-end-users": 0, "--inactive-only": 0,
                      "--active-only": 0, "--out": 1}

def parse_ts(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None

def parse_day(value: str, end: bool = False) -> datetime:
    dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    if end and len(value) == 10:      # bare YYYY-MM-DD = include that whole day
        dt += timedelta(days=1) - timedelta(seconds=1)
    return dt

def parse_args() -> argparse.Namespace:
    if len(sys.argv) > 1:
        raw, source = sys.argv[1:], "command line"
    else:
        default = os.getenv("ZENDESK_DEFAULT_ARGS", "").strip()
        try:
            raw = shlex.split(default) if default else []
        except ValueError as exc:
            sys.exit(f"ZENDESK_DEFAULT_ARGS could not be parsed (check quotes): {exc}")
        source = "ZENDESK_DEFAULT_ARGS" if default else "defaults"

    ap = argparse.ArgumentParser(description="Zendesk users -> Flexera SaaS import")
    ap.add_argument("--all-time", action="store_true", help="Full current roster (default)")
    ap.add_argument("--start", help="Window start, YYYY-MM-DD (UTC)")
    ap.add_argument("--end", help="Window end, YYYY-MM-DD (UTC). Default: now")
    ap.add_argument("--last-days", type=int, help="Trailing window, e.g. --last-days 90")
    ap.add_argument("--date-field", choices=["created", "updated"], default="created",
                    help="Which timestamp the window applies to (default: created)")
    ap.add_argument("--exclude-seat-class", default="",
                    help='Comma-separated roles to leave out, e.g. "Agente Light,Asesor"')
    ap.add_argument("--include-deleted", action="store_true",
                    help="Also send users whose active flag is False")
    ap.add_argument("--check-ticket-activity", action="store_true",
                    help="Look up each user's ticket status (snapshot CSV only, never sent to Flexera)")
    ap.add_argument("--inactivity-days", type=int, default=INACTIVITY_DAYS,
                    help=f"Ticket look-back window in days (default {INACTIVITY_DAYS})")
    ap.add_argument("--dry-run", action="store_true",
                    help="Pull from Zendesk and write the CSVs, but push nothing to Flexera")

    # Drop flags that belong to the extract script so one .env line serves both
    cleaned, ignored, i = [], [], 0
    while i < len(raw):
        flag = raw[i].split("=", 1)[0]
        if flag in EXTRACT_ONLY_FLAGS:
            takes = EXTRACT_ONLY_FLAGS[flag] and "=" not in raw[i]
            ignored.append(" ".join(raw[i:i + 1 + takes]))
            i += 1 + takes
            continue
        cleaned.append(raw[i])
        i += 1
    args = ap.parse_args(cleaned)

    print(f"Arguments ({source}): {shlex.join(cleaned) or '(none)'}")
    if ignored:
        print(f"  ignored (extract-script only): {', '.join(ignored)}")
    if args.all_time and (args.start or args.last_days):
        sys.exit("Use either --all-time or a window (--start / --last-days), not both.")
    return args

def main() -> None:
    dotenv_path = find_dotenv(usecwd=True) or os.path.join(OUTPUT_DIR, ".env")
    load_dotenv(dotenv_path=dotenv_path)   # before parse_args, for ZENDESK_DEFAULT_ARGS
    args = parse_args()

    now = datetime.now(timezone.utc)
    window: tuple[datetime, datetime] | None = None
    if args.last_days:
        window = (now - timedelta(days=args.last_days), now)
    elif args.start:
        window = (parse_day(args.start), parse_day(args.end, end=True) if args.end else now)
        if window[1] <= window[0]:
            sys.exit("--end must be after --start")
    if window:
        print(f"  window: {window[0]:%Y-%m-%d} to {window[1]:%Y-%m-%d} on {args.date_field}_at")
        print("  WARNING: only users in this window will be sent to Flexera.")
    else:
        print("  window: all time (full roster)")

    cfg = load_settings()
    base = f"https://{cfg['ZENDESK_DOMAIN']}.zendesk.com"
    session = requests.Session()
    session.auth = (f"{cfg['ZENDESK_EMAIL']}/token", cfg["ZENDESK_API_TOKEN"])
    session.headers.update({"Accept": "application/json"})

    print(f"Zendesk: {base}")
    custom_roles = fetch_custom_roles(session, base)
    print(f"  {len(custom_roles)} custom role(s)")
    raw = fetch_agents(session, base)

    by_id: dict[str, dict] = {}
    for u in raw:
        by_id[str(u.get("id"))] = u        # de-duplicate
    rows = [to_row(u, custom_roles) for u in by_id.values()]

    if ONLY_ACTIVE_USERS and not args.include_deleted:
        rows = [r for r in rows if r["active"] != "FALSE"]
    rows = [r for r in rows if r["id"]]

    if window:
        field = f"{args.date_field}_at"
        before = len(rows)
        rows = [r for r in rows
                if (ts := parse_ts(r[field])) is not None and window[0] <= ts <= window[1]]
        print(f"  {before - len(rows)} user(s) outside the window left out")

    # Roles to leave out entirely (--exclude-seat-class, plus old ZENDESK_EXCLUDE_ROLES)
    exclude_text = ",".join(x for x in (args.exclude_seat_class, cfg["ZENDESK_EXCLUDE_ROLES"]) if x)
    exclude = {x.strip().lower() for x in exclude_text.split(",") if x.strip()}
    excluded: dict[str, int] = {}
    if exclude:
        kept = []
        for r in rows:
            hit = next((v for v in (r["custom_role_name"], r["seat_class"])
                        if v and v.strip().lower() in exclude), None)
            if hit:
                excluded[hit] = excluded.get(hit, 0) + 1
            else:
                kept.append(r)
        rows = kept
        unmatched = exclude - {k.lower() for k in excluded}
        for name in sorted(unmatched):
            print(f"  NOTE: excluded role '{name}' matched no user - check the spelling")
    no_email = sum(1 for r in rows if not r["email"])

    filled = 0
    if FILL_SUBSCRIPTION_FROM_SEAT_CLASS:
        for r in rows:
            if not r["custom_role_name"]:
                r["custom_role_name"] = r["seat_class"]
                filled += 1

    rows.sort(key=lambda r: (r["seat_class"], r["name"].lower()))
    never = sum(1 for r in rows if not r["last_login_at"])
    activities = [r for r in rows if r["last_login_at"]] if DROP_BLANK_ACTIVITY_ROWS else rows

    snapshot_columns = SNAPSHOT_TEMPLATE
    if args.check_ticket_activity:
        add_ticket_status(session, base, rows, args.inactivity_days, now)
        snapshot_columns = SNAPSHOT_TEMPLATE + TICKET_COLUMNS

    write_csv(SNAPSHOT_CSV, rows, snapshot_columns, bom=True)
    write_csv(ASSIGNMENTS_CSV, rows, ASSIGNMENTS_TEMPLATE)
    write_csv(ACTIVITIES_CSV, activities, ACTIVITIES_TEMPLATE)

    print(f"\nWrote {len(rows)} assignment rows -> {ASSIGNMENTS_CSV}")
    for name, n in sorted(excluded.items()):
        print(f"  excluded role: {n} x {name}")
    if no_email:
        print(f"  {no_email} user(s) have no email (e.g. bots) - sent anyway, "
              f"expect Flexera warning(s)")
    if filled:
        print(f"  {filled} blank custom_role_name filled from seat_class")
    note = (f"{never} never-logged-in left out" if DROP_BLANK_ACTIVITY_ROWS
            else f"{never} never logged in - expect {never} Flexera warning(s)")
    print(f"Wrote {len(activities)} activity rows   -> {ACTIVITIES_CSV}  ({note})")
    print(f"Snapshot for review                -> {SNAPSHOT_CSV}"
          + ("  (includes ticket status; not sent to Flexera)" if args.check_ticket_activity else ""))

    if args.dry_run:
        print("\n--dry-run: nothing sent to Flexera.")
        return

    token = get_access_token(cfg["FLEXERA_CLIENT_ID"], cfg["FLEXERA_CLIENT_SECRET"])
    if rows:
        push_to_flexera(token, ASSIGNMENTS_CSV, ASSIGNMENTS_CONFIG_ID, "user assignments")
    if activities:
        push_to_flexera(token, ACTIVITIES_CSV, ACTIVITIES_CONFIG_ID, "user activities")
    print("\nDone. Check Flexera > Data imports for the two tasks.")

if __name__ == "__main__":
    load_dotenv(find_dotenv(usecwd=True) or os.path.join(OUTPUT_DIR, ".env"))  # for LOG_FILE
    run_logged(main, {"stdout": logging.INFO, "stderr": logging.ERROR})
