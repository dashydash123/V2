#!/usr/bin/env python3
"""
zendesk_user_extract.py
-----------------------

Pulls Zendesk users within a date window and writes the results to CSV.
Credentials and optional default command-line arguments are loaded from a
`.env` file using python-dotenv.

Setup
-----
    python -m pip install -r requirements.txt

Create a `.env` file next to this script with:
    ZENDESK_DOMAIN=your_subdomain
    ZENDESK_EMAIL=your_email@example.com
    ZENDESK_API_TOKEN=your_api_token
    ZENDESK_DEFAULT_ARGS=--last-days 90 --exclude-end-users

`ZENDESK_DOMAIN` may be either the short subdomain or a full Zendesk URL.
When the script is run without command-line arguments, it uses the arguments
stored in `ZENDESK_DEFAULT_ARGS`. Any manually supplied command-line arguments
replace the defaults completely.

Usage
-----
# Everyone created in FY26 Q1
python zendesk_user_extract.py --start 2026-04-01 --end 2026-06-30
# Trailing window, with no dates to type
python zendesk_user_extract.py --last-days 90 --exclude-end-users
# Only team members, excluding end users, filtered by last-modified date
python zendesk_user_extract.py --start 2026-01-01 --end 2026-06-30 --date-field updated --exclude-end-users
# Only roles that require paid seats
python zendesk_user_extract.py --start 2026-01-01 --end 2026-06-30 --exclude-end-users --exclude-seat-class "Light agent,Contributor"
# Run using the default arguments stored in the .env file
python zendesk_user_extract.py
# Account lifecycle: every row is labelled new / active / suspended / deactivated.
# --recent-days sets what "new" and "recently deactivated" mean (default 30)
python zendesk_user_extract.py --all-time --exclude-end-users --recent-days 30
# Subsets: new accounts only, deactivated only, or both together (new OR deactivated)
python zendesk_user_extract.py --all-time --new-only
python zendesk_user_extract.py --all-time --deactivated-only

Security
--------
The API token is used only to authenticate the HTTPS session. It is never
printed or included in validation or HTTP error messages. Keep `.env` out of
source control.
"""

from __future__ import annotations

import argparse
import csv
import logging
import os
import platform
import re
import shlex
import sys
import time
from collections import Counter
from datetime import datetime, timedelta, timezone
from logging.handlers import RotatingFileHandler

import requests
from dotenv import find_dotenv, load_dotenv

API_TIMEOUT = 60
PAGE_PAUSE = 1.5          # incremental export has its own tight rate limit
MAX_RETRIES = 5

# From the Users API reference (role_type is read-only, set by Zendesk)
ROLE_TYPE_LABELS = {
    0: "Custom agent",
    1: "Light agent",
    2: "Chat agent",
    3: "Contributor",
    4: "Admin",
    5: "Billing admin",
}

CSV_COLUMNS = [
    "id",
    "name",
    "email",
    "role",               # raw API value: end-user / agent / admin
    "role_type",          # raw integer
    "seat_class",         # human label derived from role + role_type
    "custom_role_name",   # resolved from /custom_roles
    "active",
    "suspended",
    "created_at",
    "updated_at",
    "last_login_at",
    "days_since_login",   # int, or "" if never logged in
    "inactive_90d",       # TRUE only if last login > 90 days ago (never logged in = FALSE)
    "tickets_assigned",   # populated for every user when --check-ticket-activity runs
    "tickets_submitted",
    "last_ticket_activity_at",     # most recent assigned/submitted ticket update, any date
    "days_since_ticket_activity",  # int, or "" if the user has no tickets at all
    "no_tickets_90d",     # TRUE if no assigned/submitted ticket activity in the last 90 days
    "reclaim_candidate",  # TRUE only if login dormant AND no ticket activity found
    "organization_id",
    # --- account lifecycle (for dashboard slicing) ---
    "days_since_created",       # whole days since created_at
    "is_new_account",           # TRUE if created within --recent-days (default 30)
    "account_status",           # Active / Suspended / Deactivated (Zendesk active=false = deleted)
    "is_recently_deactivated",  # TRUE if Deactivated and updated_at is within --recent-days.
                                # Zendesk gives no deactivation date; updated_at is the closest proxy.
]

# Only a recorded last login older than this counts as inactive. A user with no
# last_login_at is NOT flagged - there is no date to measure against. Kept as a
# named constant so the threshold is easy to change.
INACTIVITY_DAYS = 90

# What counts as "recent" for the account lifecycle columns: a new account was
# created within this many days; a recently deactivated one changed within it.
RECENT_DAYS = 30

SEARCH_PAUSE = 0.7        # search is rate limited harder than the users endpoint

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# --------------------------------------------------------------------------- #
# Logging
# --------------------------------------------------------------------------- #
# Every run is appended to output/<script>.log (override with LOG_FILE in .env).
# Each run starts with a separator line and every line is timestamped. Whatever
# the script prints is also written to the log, except the per-user progress
# lines. Secrets are never logged: the API token/secret values and any URL query
# string (Flexera's upload URL carries a signed token) are redacted. When the
# file passes LOG_MAX_BYTES it is renamed to <name>.log.1 and a new one starts.

LOG_MAX_BYTES = 5 * 1024 * 1024
LOG_LABEL     = "Zendesk_User_Extrsct_Env"
_SECRET_ENV   = ("ZENDESK_API_TOKEN",)
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

# --------------------------------------------------------------------------- #
# HTTP plumbing
# --------------------------------------------------------------------------- #

def build_session(email: str, token: str) -> requests.Session:
    s = requests.Session()
    s.auth = (f"{email}/token", token)
    s.headers.update({"Accept": "application/json"})
    return s

def get_json(session: requests.Session, url: str, params: dict | None = None) -> dict:
    """GET with 429 / 5xx retry, honouring Retry-After."""
    for attempt in range(1, MAX_RETRIES + 1):
        resp = session.get(url, params=params, timeout=API_TIMEOUT)

        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 60))
            print(f" rate limited, sleeping {wait}s", file=sys.stderr)
            time.sleep(wait)
            continue

        if resp.status_code in (500, 502, 503, 504) and attempt < MAX_RETRIES:
            wait = 2 ** attempt
            print(f" HTTP {resp.status_code}, retrying in {wait}s", file=sys.stderr)
            time.sleep(wait)
            continue

        if resp.status_code == 401:
            raise SystemExit(
                "401 Unauthorised. Check ZENDESK_EMAIL / ZENDESK_API_TOKEN, and that "
                "token access is enabled in Admin Center > Apps and integrations > Zendesk API."
            )
        if resp.status_code == 403:
            raise SystemExit(
                "403 Forbidden. The incremental user export endpoint requires an admin account."
            )

        resp.raise_for_status()
        return resp.json()

    raise SystemExit(f"Gave up after {MAX_RETRIES} attempts on {url}")

# --------------------------------------------------------------------------- #
# Fetching
# --------------------------------------------------------------------------- #

def fetch_custom_roles(session: requests.Session, base: str) -> dict[int, str]:
    """id -> name. Returns {} on non-Enterprise plans, which is fine."""
    try:
        data = get_json(session, f"{base}/api/v2/custom_roles")
    except requests.HTTPError:
        print(" custom roles unavailable on this plan, continuing", file=sys.stderr)
        return {}
    roles = {r["id"]: r.get("name", "") for r in data.get("custom_roles", [])}
    print(f" resolved {len(roles)} custom role(s)", file=sys.stderr)
    return roles

def fetch_users_since(session: requests.Session, base: str, start_epoch: int) -> list[dict]:
    """Cursor-based incremental export from start_epoch to now."""
    url = f"{base}/api/v2/incremental/users/cursor"
    params = {"start_time": start_epoch, "per_page": 1000}
    users: list[dict] = []
    page = 0

    while True:
        page += 1
        data = get_json(session, url, params)
        batch = data.get("users", [])
        users.extend(batch)
        print(f" page {page}: +{len(batch)} (running total {len(users)})", file=sys.stderr)

        if data.get("end_of_stream"):
            break
        cursor = data.get("after_cursor")
        if not cursor:
            break
        params = {"cursor": cursor}  # start_time is only for the first request
        time.sleep(PAGE_PAUSE)

    return users

def fetch_users_list(session: requests.Session, base: str, roles: list[str] | None = None) -> list[dict]:
    """
    Plain GET /api/v2/users with cursor pagination (page[size]/page[after]).
    Unlike the incremental export, this endpoint accepts role[]=agent&role[]=admin
    and filters server-side - Zendesk never sends the end-user rows at all.
    Use this whenever there's no date window to honour (i.e. --all-time).
    """
    url = f"{base}/api/v2/users"
    params: dict | None = {"page[size]": 100}
    if roles:
        params["role[]"] = roles  # requests encodes a list as repeated role[]=a&role[]=b
    users: list[dict] = []
    page = 0

    while url:
        page += 1
        data = get_json(session, url, params)
        batch = data.get("users", [])
        users.extend(batch)
        print(f" page {page}: +{len(batch)} (running total {len(users)})", file=sys.stderr)

        meta = data.get("meta", {})
        if not meta.get("has_more"):
            break
        url = (data.get("links") or {}).get("next")
        params = None  # the next link already has page[size]/page[after]/role[] encoded
        time.sleep(0.3)

    return users

def search_latest(session: requests.Session, base: str, query: str) -> tuple[int, datetime | None]:
    """
    One /api/v2/search call returning (total_count, updated_at of newest match).
    per_page=1 sorted newest-first gives both the count and the latest date
    for the price of a single request.
    """
    data = get_json(session, f"{base}/api/v2/search.json", {
        "query": query,
        "sort_by": "updated_at",
        "sort_order": "desc",
        "per_page": 1,
    })
    results = data.get("results") or []
    latest = parse_ts(results[0].get("updated_at")) if results else None
    return int(data.get("count", 0)), latest

def _newest(*stamps: datetime | None) -> datetime | None:
    present = [t for t in stamps if t is not None]
    return max(present) if present else None

def ticket_activity(session: requests.Session, base: str, user_id: int,
                    since: datetime) -> tuple[int, int, datetime | None]:
    """
    Return (tickets_assigned, tickets_submitted, last_ticket_activity).

    The two counts cover tickets touched since `since`. last_ticket_activity is
    the newest updated_at across assigned and submitted tickets with no date
    limit, so a user idle for 200 days shows 200 rather than a blank. The
    unbounded lookup only runs (2 extra calls) when the window found nothing.

    Zendesk ticket search supports the assignee, submitter and requester
    keywords only - there is no 'commenter' keyword. A Light Agent who only
    ever adds private comments will therefore score zero here even if they
    are working daily. Treat a zero as "needs a human to confirm", not as
    proof the seat is unused.
    """
    day = since.strftime("%Y-%m-%d")
    assigned, last_a = search_latest(session, base, f"type:ticket assignee:{user_id} updated>{day}")
    time.sleep(SEARCH_PAUSE)
    submitted, last_s = search_latest(session, base, f"type:ticket submitter:{user_id} updated>{day}")
    time.sleep(SEARCH_PAUSE)
    last = _newest(last_a, last_s)

    if last is None:
        # Nothing in the window - look back without a date filter to find when
        # (if ever) this user last had a ticket.
        _, last_a = search_latest(session, base, f"type:ticket assignee:{user_id}")
        time.sleep(SEARCH_PAUSE)
        _, last_s = search_latest(session, base, f"type:ticket submitter:{user_id}")
        time.sleep(SEARCH_PAUSE)
        last = _newest(last_a, last_s)

    return assigned, submitted, last

# --------------------------------------------------------------------------- #
# Transform
# --------------------------------------------------------------------------- #

def classify(user: dict, custom_roles: dict[int, str]) -> tuple[str, str]:
    """Return (seat_class, custom_role_name)."""
    role = user.get("role") or ""
    role_type = user.get("role_type")
    crid = user.get("custom_role_id")
    custom_name = custom_roles.get(crid, "") if crid else ""

    if role == "end-user":
        return "End user", ""

    if role_type in ROLE_TYPE_LABELS:
        label = ROLE_TYPE_LABELS[role_type]
        # A "custom agent" is only meaningful with its role name attached
        if role_type == 0 and custom_name:
            label = f"Custom agent: {custom_name}"
        return label, custom_name

    if role == "admin":
        return "Admin", custom_name
    if role == "agent":
        return "Agent", custom_name
    return role or "Unknown", custom_name

def parse_ts(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None

def login_recency(last_login_at: str | None, now: datetime, threshold: int) -> tuple[str, str]:
    """
    Return (days_since_login, inactive_flag).

    A null last_login_at is reported as a blank day count and is NOT flagged as
    inactive: with no login date there is nothing to prove it is older than
    the threshold. These users are still counted separately in the summary.
    """
    ts = parse_ts(last_login_at)
    if ts is None:
        return "", "FALSE"
    days = (now - ts).days
    return str(days), "TRUE" if days > threshold else "FALSE"

def account_status(user: dict) -> str:
    """Deactivated = Zendesk active is false (deleted user); Suspended = suspended is true."""
    if user.get("active") is False:
        return "Deactivated"
    if user.get("suspended") is True:
        return "Suspended"
    return "Active"

def account_lifecycle(user: dict, status: str, now: datetime, recent_days: int) -> tuple[str, str, str]:
    """Return (days_since_created, is_new_account, is_recently_deactivated)."""
    created = parse_ts(user.get("created_at"))
    updated = parse_ts(user.get("updated_at"))
    days_created = (now - created).days if created else ""
    is_new = "TRUE" if created and (now - created).days < recent_days else "FALSE"
    recently_deact = ("TRUE" if status == "Deactivated" and updated
                      and (now - updated).days < recent_days else "FALSE")
    return days_created, is_new, recently_deact

def to_row(user: dict, custom_roles: dict[int, str], now: datetime, threshold: int,
           recent_days: int = RECENT_DAYS) -> dict:
    seat_class, custom_name = classify(user, custom_roles)
    days_since_login, inactive = login_recency(user.get("last_login_at"), now, threshold)
    status = account_status(user)
    days_created, is_new, recently_deact = account_lifecycle(user, status, now, recent_days)
    return {
        "id": user.get("id"),
        "name": user.get("name") or "",
        "email": user.get("email") or "",
        "role": user.get("role") or "",
        "role_type": user.get("role_type") if user.get("role_type") is not None else "",
        "seat_class": seat_class,
        "custom_role_name": custom_name,
        "active": user.get("active"),
        "suspended": user.get("suspended"),
        "created_at": user.get("created_at") or "",
        "updated_at": user.get("updated_at") or "",
        "last_login_at": user.get("last_login_at") or "",
        "days_since_login": days_since_login,
        "inactive_90d": inactive,
        "tickets_assigned": "",
        "tickets_submitted": "",
        "last_ticket_activity_at": "",
        "days_since_ticket_activity": "",
        "no_tickets_90d": "",
        "reclaim_candidate": "",
        "organization_id": user.get("organization_id") or "",
        "days_since_created": days_created,
        "is_new_account": is_new,
        "account_status": status,
        "is_recently_deactivated": recently_deact,
    }

# --------------------------------------------------------------------------- #
# Environment configuration
# --------------------------------------------------------------------------- #

def sanitize_domain(raw: str) -> str:
    """Return the short Zendesk subdomain from a short name or full URL."""
    value = raw.strip()
    if "://" in value:
        value = value.split("://", 1)[1]
    value = value.strip("/")
    value = value.split("/")[0]
    if value.endswith(".zendesk.com"):
        value = value[:-len(".zendesk.com")]
    return value.strip()

def load_environment() -> tuple[str, str, str]:
    """Load and validate required Zendesk settings without exposing secrets."""
    # Prefer a .env in the current working directory. If none is found, search
    # upward from the current directory, which also works in Windows Command Prompt.
    dotenv_path = find_dotenv(usecwd=True)
    load_dotenv(dotenv_path=dotenv_path or None)

    required = {
        "ZENDESK_DOMAIN": os.getenv("ZENDESK_DOMAIN", "").strip(),
        "ZENDESK_EMAIL": os.getenv("ZENDESK_EMAIL", "").strip(),
        "ZENDESK_API_TOKEN": os.getenv("ZENDESK_API_TOKEN", "").strip(),
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        location = f" at {dotenv_path}" if dotenv_path else ""
        raise SystemExit(
            "Missing or empty required environment variable(s): "
            + ", ".join(missing)
            + f". Create or update the .env file{location}."
        )

    domain = sanitize_domain(required["ZENDESK_DOMAIN"])
    if not domain or "." in domain:
        raise SystemExit(
            "ZENDESK_DOMAIN is invalid. Use only the Zendesk subdomain "
            " (for example, 'yourcompany') or a URL such as "
            " 'https://yourcompany.zendesk.com'."
        )

    return domain, required["ZENDESK_EMAIL"], required["ZENDESK_API_TOKEN"]

def command_line_args() -> list[str]:
    """Return manual CLI args, or safely split ZENDESK_DEFAULT_ARGS if absent."""
    if len(sys.argv) > 1:
        return sys.argv[1:]

    default_args = os.getenv("ZENDESK_DEFAULT_ARGS", "").strip()
    if not default_args:
        return []

    try:
        # shlex.split removes grouping quotes while preserving the quoted value
        # as one argument, including on Windows Command Prompt.
        return shlex.split(default_args)
    except ValueError as exc:
        raise SystemExit(
            "ZENDESK_DEFAULT_ARGS could not be parsed. Check that quoted values "
            f"have matching quotation marks. Details: {exc}"
        ) from exc

# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def parse_window(start: str, end: str | None) -> tuple[datetime, datetime]:
    start_dt = datetime.fromisoformat(start)
    if start_dt.tzinfo is None:
        start_dt = start_dt.replace(tzinfo=timezone.utc)
    if end:
        end_dt = datetime.fromisoformat(end)
        if end_dt.tzinfo is None:
            end_dt = end_dt.replace(tzinfo=timezone.utc)
        if len(end) == 10:  # bare YYYY-MM-DD means include the whole day
            end_dt += timedelta(days=1) - timedelta(seconds=1)
    else:
        end_dt = datetime.now(timezone.utc)
    if end_dt <= start_dt:
        raise SystemExit("--end must be after --start")
    return start_dt, end_dt

def main() -> None:
    # Load .env before parsing so it can supply default command-line arguments.
    load_dotenv(find_dotenv(usecwd=True) or None)

    p = argparse.ArgumentParser(description="Export Zendesk users (name, email, role) for a date window.")
    p.add_argument("--start", help="Window start, YYYY-MM-DD or ISO datetime (UTC)")
    p.add_argument("--end", help="Window end, YYYY-MM-DD or ISO datetime (UTC). Default: now")
    p.add_argument("--last-days", type=int,
                   help="Shorthand for a trailing window, e.g. --last-days 90. Alternative to --start")
    p.add_argument("--all-time", action="store_true",
                   help="Fetch every user regardless of created/updated date - the full current roster. "
                        "Overrides --start/--last-days/--date-field.")
    p.add_argument("--date-field", choices=["created", "updated"], default="created",
                   help="Which timestamp the window applies to (default: created)")
    p.add_argument("--out", default="zendesk_users.csv", help="Output CSV path")
    p.add_argument("--exclude-end-users", action="store_true",
                   help="Drop role=end-user, leaving team members only")
    p.add_argument("--exclude-seat-class", default="",
                   help='Comma-separated seat_class values to drop, e.g. "Light agent,Contributor"')
    p.add_argument("--active-only", action="store_true", help="Drop deleted/inactive users")
    p.add_argument("--inactivity-days", type=int, default=INACTIVITY_DAYS,
                   help=f"Days without a login before a seat counts as dormant (default: {INACTIVITY_DAYS})")
    p.add_argument("--inactive-only", action="store_true",
                   help="Keep only dormant seats - the reclaim candidate list")
    p.add_argument("--recent-days", type=int, default=RECENT_DAYS,
                   help=f"What counts as recent for is_new_account and is_recently_deactivated (default: {RECENT_DAYS})")
    p.add_argument("--new-only", action="store_true",
                   help="Keep only accounts created within --recent-days")
    p.add_argument("--deactivated-only", action="store_true",
                   help="Keep only deactivated accounts (active = false). "
                        "With --new-only as well, keeps accounts that are new OR deactivated")
    p.add_argument("--check-ticket-activity", action="store_true",
                   help="For every exported user, query the Search API for assigned/submitted ticket "
                        "activity in the inactivity window, plus the date of their last ticket. "
                        "2 API calls per user, plus 2 more for users with nothing in the window. "
                        "Fills the ticket columns, no_tickets_90d and reclaim_candidate.")
    args = p.parse_args(command_line_args())
    if args.active_only and args.deactivated_only:
        raise SystemExit("--active-only and --deactivated-only contradict each other: "
                         "--active-only drops exactly the accounts --deactivated-only keeps.")

    if not args.start and not args.last_days and not args.all_time:
        raise SystemExit("Give --start YYYY-MM-DD, --last-days N, or --all-time")
    if args.all_time:
        # Zendesk didn't exist before 2007; this just means "since forever".
        start_dt = datetime(2007, 1, 1, tzinfo=timezone.utc)
        end_dt = datetime.now(timezone.utc)
    elif args.last_days:
        start_dt = datetime.now(timezone.utc) - timedelta(days=args.last_days)
        end_dt = datetime.now(timezone.utc)
    else:
        start_dt, end_dt = parse_window(args.start, args.end)

    print(f"Arguments: {shlex.join(command_line_args()) or '(none)'}", file=sys.stderr)
    print("Resolving credentials...", file=sys.stderr)
    subdomain, email, token = load_environment()
    base = f"https://{subdomain}.zendesk.com"
    print(f" target: {base}", file=sys.stderr)
    session = build_session(email, token)

    print(f"Window: {start_dt:%Y-%m-%d %H:%M} to {end_dt:%Y-%m-%d %H:%M} UTC "
          f"on {args.date_field}_at", file=sys.stderr)

    print("Fetching custom roles...", file=sys.stderr)
    custom_roles = fetch_custom_roles(session, base)

    if args.all_time:
        role_filter = ["agent", "admin"] if args.exclude_end_users else None
        label = "roles: agent, admin (end users filtered server-side)" if role_filter else "all roles"
        print(f"Fetching users (full list, {label})...", file=sys.stderr)
        raw_users = fetch_users_list(session, base, roles=role_filter)
    else:
        print("Fetching users (incremental export)...", file=sys.stderr)
        raw_users = fetch_users_since(session, base, int(start_dt.timestamp()))

    # Deduplicate: time-based/cursor exports can repeat a record across pages
    seen: dict[int, dict] = {}
    for u in raw_users:
        seen[u["id"]] = u

    field = f"{args.date_field}_at"
    # Matched case-insensitively against seat_class OR custom_role_name, so
    # "Agente Light" and "Light agent" both work.
    drop_classes = {c.strip().lower() for c in args.exclude_seat_class.split(",") if c.strip()}
    # Pinned once so every row is measured against the same instant, not a
    # clock that drifts forward while the loop runs.
    now_utc = datetime.now(timezone.utc)

    rows = []
    for u in seen.values():
        ts = parse_ts(u.get(field))
        if ts is None or not (start_dt <= ts <= end_dt):
            continue

        if args.exclude_end_users and u.get("role") == "end-user":
            continue

        if args.active_only and u.get("active") is False:
            continue

        row = to_row(u, custom_roles, now_utc, args.inactivity_days, args.recent_days)
        if (row["seat_class"].strip().lower() in drop_classes
                or row["custom_role_name"].strip().lower() in drop_classes):
            continue

        if args.inactive_only and row["inactive_90d"] != "TRUE":
            continue

        if args.new_only or args.deactivated_only:
            wanted = ((args.new_only and row["is_new_account"] == "TRUE")
                      or (args.deactivated_only and row["account_status"] == "Deactivated"))
            if not wanted:
                continue
        rows.append(row)

    if args.check_ticket_activity:
        cutoff = now_utc - timedelta(days=args.inactivity_days)
        print(f"\nChecking ticket activity for all {len(rows)} user(s) since {cutoff:%Y-%m-%d} "
              f"(at least {len(rows) * 2} API calls)...", file=sys.stderr)
        for i, r in enumerate(rows, 1):
            assigned, submitted, last = ticket_activity(session, base, r["id"], cutoff)
            no_tickets = (assigned + submitted) == 0
            r["tickets_assigned"] = assigned
            r["tickets_submitted"] = submitted
            if last is not None:
                r["last_ticket_activity_at"] = last.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                r["days_since_ticket_activity"] = (now_utc - last).days
            r["no_tickets_90d"] = "TRUE" if no_tickets else "FALSE"
            # A deactivated (deleted) account holds no seat, so it is never a reclaim candidate
            r["reclaim_candidate"] = ("TRUE" if (r["inactive_90d"] == "TRUE" and no_tickets
                                                 and r["account_status"] != "Deactivated") else "FALSE")
            if not no_tickets:
                flag = f"{assigned + submitted} ticket(s)"
            elif last is not None:
                flag = f"no activity, last ticket {r['days_since_ticket_activity']}d ago"
            else:
                flag = "no tickets ever"
            print(f"  [{i}/{len(rows)}] {r['name'][:34]:<34} {flag}", file=sys.stderr)

    rows.sort(key=lambda r: (r["seat_class"], r["name"].lower()))

    with open(args.out, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nFetched {len(seen)} unique users, {len(rows)} in window -> {args.out}", file=sys.stderr)
    print("\nBreakdown by seat_class:", file=sys.stderr)
    dormant_by_class = Counter(r["seat_class"] for r in rows if r["inactive_90d"] == "TRUE")
    for cls, n in Counter(r["seat_class"] for r in rows).most_common():
        print(f"  {cls:<32} {n:>6}   ({dormant_by_class[cls]} dormant)", file=sys.stderr)

    dormant = sum(dormant_by_class.values())
    never = sum(1 for r in rows if not r["days_since_login"])
    print(f"\nDormant (last login more than {args.inactivity_days} days ago): {dormant} of {len(rows)}",
          file=sys.stderr)
    print(f"No last login date recorded (not counted as dormant): {never}", file=sys.stderr)

    new_n = sum(1 for r in rows if r["is_new_account"] == "TRUE")
    deact_n = sum(1 for r in rows if r["account_status"] == "Deactivated")
    recent_deact_n = sum(1 for r in rows if r["is_recently_deactivated"] == "TRUE")
    susp_n = sum(1 for r in rows if r["account_status"] == "Suspended")
    print(f"\nAccount lifecycle (recent = last {args.recent_days} days):", file=sys.stderr)
    print(f"  new accounts created in the last {args.recent_days} days: {new_n}", file=sys.stderr)
    print(f"  deactivated (active = false): {deact_n}   "
          f"({recent_deact_n} changed in the last {args.recent_days} days, by updated_at)", file=sys.stderr)
    print(f"  suspended: {susp_n}", file=sys.stderr)
    if deact_n == 0 and not (args.active_only or args.new_only):
        print("  Note: no deactivated accounts in this extract. Zendesk may list deleted users "
              "separately, and this script does not read that list.", file=sys.stderr)

    if args.check_ticket_activity:
        no_tix = sum(1 for r in rows if r["no_tickets_90d"] == "TRUE")
        no_tix_ever = sum(1 for r in rows if r["no_tickets_90d"] == "TRUE"
                          and r["days_since_ticket_activity"] == "")
        confirmed = sum(1 for r in rows if r["reclaim_candidate"] == "TRUE")
        contradicted = sum(1 for r in rows if r["inactive_90d"] == "TRUE"
                           and r["no_tickets_90d"] == "FALSE")
        print(f"\nNo ticket activity in {args.inactivity_days} days: {no_tix} of {len(rows)}"
              f"  [{no_tix_ever} have no tickets at all]", file=sys.stderr)
        print(f"  dormant login AND no ticket activity (reclaim candidates): {confirmed}", file=sys.stderr)
        print(f"  dormant login but WORKING tickets: {contradicted} <- do not reclaim these",
              file=sys.stderr)
        print(f"\nNote: search covers assignee and submitter only. Light Agents who just "
              f"comment will show zero. Confirm with role owners before removing anyone.",
              file=sys.stderr)

if __name__ == "__main__":
    load_dotenv(find_dotenv(usecwd=True) or None)   # so LOG_FILE can come from .env
    run_logged(main, {"stderr": logging.INFO})
