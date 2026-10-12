#!/usr/bin/env python3
"""anaplan_flexera_sync.py - Anaplan users -> Flexera (Snow Atlas) SaaS import, one run.

  1. Anaplan  authenticate (saved refresh token or device-code), run the
              export, download every chunk -> test.csv, post-process it
              (A1 metadata row, Email ID header, ISO dates, inactivity columns)
  2. Build    anaplan_license.csv    : Email ID, User ID, License Type,
                                       Account Creation Date, Last Active Date
              anaplan_activities.csv : User ID, Last Active Date
              License Type = <License Type> + separator + <Assigned Zone>
              (zone normalised to GHQ/NAZ/MAZ/SAZ/APAC/EUR/AFR or kept as-is,
              blank zones -> ZONE_BLANK_DEFAULT; all set in .env)
  3. Push     both CSVs: token -> upload URL -> PUT to blob -> trigger import

All IDs, keys, URLs and switches live in .env in the same folder as this
script. If .env is missing, a blank template is created and the run stops.
Every run is written to logs/anaplan_flexera_sync.log (and the console).
Output files keep the same names and are overwritten on each run.

Requires:  python -m pip install requests pandas python-dotenv
"""

import csv
import hashlib
import json
import logging
import os
import sys
import time
from datetime import date, datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path

import pandas as pd
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
ENV_FILE = SCRIPT_DIR / ".env"

# ===========================================================================
# .ENV - created from this template when it does not exist yet
# ===========================================================================

ENV_TEMPLATE = r'''# Anaplan -> Flexera (Snow Atlas) SaaS import - settings for anaplan_flexera_sync.py
# Keep this file next to the script and named exactly ".env".
# Values with spaces or special characters (& ^ @ # ;) go in single quotes.

# ═════════════ Anaplan ═════════════
# Anaplan OAuth client ID (NOT the Flexera one)
ANAPLAN_CLIENT_ID=
ANAPLAN_WORKSPACE_ID=
ANAPLAN_MODEL_ID=
ANAPLAN_EXPORT_ID=
ANAPLAN_SCOPE='openid profile email offline_access'
ANAPLAN_DEVICE_CODE_URL=https://us1a.app.anaplan.com/oauth/device/code
ANAPLAN_TOKEN_URL=https://us1a.app.anaplan.com/oauth/token
ANAPLAN_API_BASE_URL=https://api.anaplan.com/2/0

# ═════════════ Flexera (Snow Atlas) SaaS import - UAT ═════════════
# Application Registration with 'Manage SaaS data import API integrations'
FLEXERA_CLIENT_ID=
FLEXERA_CLIENT_SECRET=

FLEXERA_BASE_URL=https://snowatlas-westeurope.flexera.eu
# Leave blank to build them from FLEXERA_BASE_URL
FLEXERA_TOKEN_URL=
FLEXERA_UPLOAD_URL=

# Anaplan import configs (Flexera > Data imports > Anaplan > API Integration)
#   User assignment: Email ID -> Email address | User ID -> Unique user identifier |
#                    License Type -> Subscription name | Account Creation Date ->
#                    User account creation | Last Active Date -> User account last updated
FLEXERA_ASSIGNMENTS_CONFIG_ID=3298061a-7c0b-4218-bb95-739f249e3a56
#   User activities: User ID -> Unique user identifier |
#                    Last Active Date -> Activity name + recorded value
FLEXERA_ACTIVITIES_CONFIG_ID=99429703-8c03-4980-ac89-fb830fa855bb

# ═════════════ Run switches (yes / no) ═════════════
# no = reuse the existing export file instead of calling Anaplan
RUN_ANAPLAN_EXPORT=yes
# no = build the CSVs only, nothing sent to Flexera
FLEXERA_PUSH=yes
# yes = drop users whose User State is not 'Enabled'
ONLY_ENABLED_USERS=yes

# ═════════════ Assigned Zone -> License Type ═════════════
# License Type sent to Flexera = <License Type> + separator + <zone>
#   e.g. "Enterprise User" + " " + "SAZ" -> "Enterprise User SAZ"
LICENSE_ZONE_SEPARATOR=' '

# yes = map every zone to one of GHQ, NAZ, MAZ, SAZ, APAC, EUR, AFR
# no  = keep the Assigned Zone exactly as it is in Anaplan
ZONE_NORMALISE=yes

# Zone used when Assigned Zone is blank - type any value (used with yes AND no)
ZONE_BLANK_DEFAULT=GHQ

# Only when ZONE_NORMALISE=yes - zone used when a value is not recognised.
# Type a zone (e.g. GHQ), or KEEP to leave the original text. Logged as a warning.
ZONE_UNMAPPED_DEFAULT=GHQ

# Only when ZONE_NORMALISE=yes - extra spellings without editing the script.
# Format: variant=ZONE;variant=ZONE   e.g. 'Brasil=SAZ;Corporate Affairs=GHQ'
ZONE_EXTRA_ALIASES=

# ═════════════ Files ═════════════
# Folder for token.json and all CSVs (blank = this script's folder)
OUTPUT_DIR=
# Same names on every run -> files are overwritten, not duplicated
ANAPLAN_EXPORT_FILE=test.csv
ASSIGNMENTS_FILE=anaplan_license.csv
ACTIVITIES_FILE=anaplan_activities.csv

# ═════════════ Log ═════════════
# Blank = <script folder>\logs
LOG_DIR=
LOG_FILE=anaplan_flexera_sync.log
# New log file when it reaches LOG_MAX_MB; LOG_BACKUPS old files are kept
LOG_MAX_MB=5
LOG_BACKUPS=10
# DEBUG / INFO / WARNING
LOG_LEVEL=INFO
'''


def _load_env_file(path: Path) -> None:
    """Load KEY=VALUE lines into os.environ (python-dotenv if installed)."""
    try:
        from dotenv import load_dotenv
        load_dotenv(path, override=True)
        return
    except ImportError:
        pass
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "'\"":
            value = value[1:-1]
        else:
            value = value.split(" #", 1)[0].strip()
        os.environ[key.strip()] = value


if not ENV_FILE.exists():
    ENV_FILE.write_text(ENV_TEMPLATE, encoding="utf-8")
    sys.exit(f"[config] No .env found - a blank template was created at:\n  {ENV_FILE}\n"
             "Fill in the Anaplan and Flexera values, then run the script again.")
_load_env_file(ENV_FILE)


def env(key: str, default: str = "") -> str:
    """Read a setting; stray spaces and quotes from pasting are removed."""
    value = os.getenv(key)
    if value is None:
        return default
    value = value.strip().strip('"').strip("'")
    return value if value else default


def env_raw(key: str, default: str = "") -> str:
    """Like env() but keeps spaces (used for separators)."""
    value = os.getenv(key)
    return default if value is None else value


def env_bool(key: str, default: bool) -> bool:
    value = env(key, "yes" if default else "no").lower()
    if value in ("yes", "y", "true", "1", "on"):
        return True
    if value in ("no", "n", "false", "0", "off"):
        return False
    sys.exit(f"[config] {key}={value!r} in .env - use yes or no")


def env_path(key: str, default: Path) -> Path:
    value = env(key)
    if not value:
        return default
    path = Path(value)
    return path if path.is_absolute() else SCRIPT_DIR / path


# ===========================================================================
# CONFIGURATION (values come from .env)
# ===========================================================================

# ---- Anaplan --------------------------------------------------------------
WORKSPACE_ID      = env("ANAPLAN_WORKSPACE_ID")
MODEL_ID          = env("ANAPLAN_MODEL_ID")
EXPORT_ID         = env("ANAPLAN_EXPORT_ID")
ANAPLAN_CLIENT_ID = env("ANAPLAN_CLIENT_ID")
SCOPE = env("ANAPLAN_SCOPE", "openid profile email offline_access")

OAUTH_DEVICE_CODE_URL = env("ANAPLAN_DEVICE_CODE_URL", "https://us1a.app.anaplan.com/oauth/device/code")
OAUTH_TOKEN_URL       = env("ANAPLAN_TOKEN_URL", "https://us1a.app.anaplan.com/oauth/token")
ANAPLAN_API_BASE_URL  = env("ANAPLAN_API_BASE_URL", "https://api.anaplan.com/2/0").rstrip("/")

# ---- Flexera (UAT) - Application Registration -----------------------------
CLIENT_ID     = env("FLEXERA_CLIENT_ID")
CLIENT_SECRET = env("FLEXERA_CLIENT_SECRET")

BASE        = env("FLEXERA_BASE_URL", "https://snowatlas-westeurope.flexera.eu").rstrip("/")
TOKEN_URL   = env("FLEXERA_TOKEN_URL", f"{BASE}/idp/api/connect/token")
USERS_URL   = f"{BASE}/api/saas/consolidated-view/v1/users"
UPLOAD_URL  = env("FLEXERA_UPLOAD_URL", f"{BASE}/api/saas/import/v1/uploads")
# IMPORT_URL = f"{BASE}/api/saas/import/v1/configs/{{CONFIG_ID}}/records/import"

ASSIGNMENTS_CONFIG_ID = env("FLEXERA_ASSIGNMENTS_CONFIG_ID")   # user assignments
ACTIVITIES_CONFIG_ID  = env("FLEXERA_ACTIVITIES_CONFIG_ID")    # user activities

# ---- Files (fixed names -> overwritten every run) -------------------------
BASE_DIRECTORY  = env_path("OUTPUT_DIR", SCRIPT_DIR)
TOKEN_FILE      = BASE_DIRECTORY / "token.json"
OUTPUT_FILE     = BASE_DIRECTORY / env("ANAPLAN_EXPORT_FILE", "test.csv")
ASSIGNMENTS_CSV = BASE_DIRECTORY / env("ASSIGNMENTS_FILE", "anaplan_license.csv")
ACTIVITIES_CSV  = BASE_DIRECTORY / env("ACTIVITIES_FILE", "anaplan_activities.csv")

LOG_DIR     = env_path("LOG_DIR", SCRIPT_DIR / "logs")
LOG_PATH    = LOG_DIR / env("LOG_FILE", "anaplan_flexera_sync.log")
LOG_MAX_MB  = env("LOG_MAX_MB", "5")
LOG_BACKUPS = env("LOG_BACKUPS", "10")
LOG_LEVEL   = env("LOG_LEVEL", "INFO").upper()

# ---- Run switches ---------------------------------------------------------
RUN_ANAPLAN_EXPORT  = env_bool("RUN_ANAPLAN_EXPORT", True)
PUSH_TO_FLEXERA     = env_bool("FLEXERA_PUSH", True)
ONLY_ENABLED_USERS  = env_bool("ONLY_ENABLED_USERS", True)
ANAPLAN_TENANT_KEY  = "anaplan"   # make_user_id() key when User ID is blank

# ---- Anaplan CSV post-processing ------------------------------------------
LAST_LOGIN_COLUMN    = "Last Active Date"
INACTIVE_DAYS_COLUMN = "Inactive Days"
INACTIVE_COLUMN      = "Not Logged In Over 90 Days"
INACTIVE_DAYS        = 90
ISO_DATETIME_COLUMNS = ("Account Creation Date", "Last Active Date")

# ---- Anaplan export header -> internal name --------------------------------
SOURCE_COLUMNS = {
    "Email ID":              "email",
    "User ID":               "anaplan_id",
    "License Type":          "license_type",
    "User State":            "user_state",
    "Assigned Zone":         "zone",
    "Account Creation Date": "created",
    "Last Active Date":      "last_active",
}
REQUIRED_SOURCE = ["email", "license_type", "last_active"]

# ---- Assigned Zone -> License Type ----------------------------------------
LICENSE_ZONE_SEPARATOR = env_raw("LICENSE_ZONE_SEPARATOR", " ").strip("'\"") or " "
ZONE_NORMALISE         = env_bool("ZONE_NORMALISE", True)
ZONE_BLANK_DEFAULT     = env("ZONE_BLANK_DEFAULT", "GHQ")
ZONE_UNMAPPED_DEFAULT  = env("ZONE_UNMAPPED_DEFAULT", "GHQ")   # or KEEP
CANONICAL_ZONES = ("GHQ", "NAZ", "MAZ", "SAZ", "APAC", "EUR", "AFR")

# Alias -> code. Matching is case-, space- and punctuation-insensitive and
# works on whole words, so "GHQ SOLUTIONS", "ghq-fp&a", "Global HQ" -> GHQ.
# Extra spellings can also be added in .env (ZONE_EXTRA_ALIASES).
ZONE_ALIASES = {
    # GHQ
    "GHQ": "GHQ", "HQ": "GHQ", "GLOBAL": "GHQ", "GLOBAL HQ": "GHQ",
    "GLOBAL HEADQUARTERS": "GHQ", "GLOBAL HEAD QUARTERS": "GHQ",
    "HEADQUARTERS": "GHQ", "CORPORATE": "GHQ", "GCC": "GHQ",
    # NAZ
    "NAZ": "NAZ", "NORTH AMERICA": "NAZ", "NORTH AMERICAN": "NAZ",
    "NORTH AMERICA ZONE": "NAZ", "NA ZONE": "NAZ", "USA": "NAZ", "US": "NAZ",
    "CANADA": "NAZ",
    # MAZ
    "MAZ": "MAZ", "MIDDLE AMERICA": "MAZ", "MIDDLE AMERICAS": "MAZ",
    "MIDDLE AMERICAS ZONE": "MAZ", "MIDDLE AMERICA ZONE": "MAZ",
    "CENTRAL AMERICA": "MAZ", "MEXICO": "MAZ",
    # SAZ (South America)
    "SAZ": "SAZ", "SOUTH AMERICA": "SAZ", "SOUTH AMERICAN": "SAZ",
    "SOUTH AMERICA ZONE": "SAZ", "SA ZONE": "SAZ", "BRAZIL": "SAZ",
    # APAC (South Asia / India included)
    "APAC": "APAC", "ASIA PACIFIC": "APAC", "ASIA PAC": "APAC", "ASPAC": "APAC",
    "AP": "APAC", "ASIA": "APAC", "APAC NORTH": "APAC", "APAC SOUTH": "APAC",
    "CHINA": "APAC",
    # EUR
    "EUR": "EUR", "EU": "EUR", "EURO": "EUR", "EUROPE": "EUR",
    "EUROPE ZONE": "EUR", "EUROPEAN": "EUR",
    # AFR
    "AFR": "AFR", "AF": "AFR", "AFRICA": "AFR", "AFRICA ZONE": "AFR",
    "AFRICAN": "AFR",
}
BLANK_ZONE_VALUES = {"", "NA", "N A", "NONE", "NULL", "NAN", "BLANK", "UNKNOWN"}

# ---- Flexera CSV headers (must match each import's Mapping screen) ---------
# Assignments: Email address = Email ID | Unique user identifier = User ID |
#   Subscription name = License Type | User account creation = Account
#   Creation Date | User account last updated = Last Active Date
ASSIGNMENTS_TEMPLATE = ["Email ID", "User ID", "License Type",
                        "Account Creation Date", "Last Active Date"]
# Activities: Unique user identifier = User ID | Activity or event name =
#   Last Active Date | Activity recorded value = Last Active Date
ACTIVITIES_TEMPLATE = ["User ID", "Last Active Date"]


# ===========================================================================
# LOGGING - console + rotating log file
# ===========================================================================

LOG = logging.getLogger("anaplan_flexera_sync")
_LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"


def setup_logging() -> None:
    if LOG.handlers:
        return
    LOG.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))
    LOG.propagate = False
    formatter = logging.Formatter(_LOG_FORMAT, "%Y-%m-%d %H:%M:%S")
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    LOG.addHandler(console)
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            LOG_PATH, maxBytes=int(float(LOG_MAX_MB) * 1024 * 1024),
            backupCount=int(LOG_BACKUPS), encoding="utf-8")
        file_handler.setFormatter(formatter)
        LOG.addHandler(file_handler)
    except (OSError, ValueError) as exc:
        LOG.warning(f"Log file could not be opened ({exc}) - logging to console only.")


def print(*args, sep=" ", end="\n", file=None, flush=False):  # noqa: A001
    """Send every message in this script to the log (console + file).
    Also applies "%s"-style arguments, so the Flexera functions' output
    ('fileId=%s', 'HTTP %s') shows the real values in the log."""
    if not LOG.handlers:
        setup_logging()
    if len(args) > 1 and isinstance(args[0], str) and "%" in args[0]:
        try:
            message = args[0] % tuple(args[1:])
        except (TypeError, ValueError):
            message = sep.join(str(a) for a in args)
    else:
        message = sep.join(str(a) for a in args)
    for line in message.strip("\n").splitlines() or [""]:
        if not line.strip():
            continue
        if line.lstrip().lower().startswith(("warning", "refresh failed")):
            LOG.warning(line)
        else:
            LOG.info(line)


# ===========================================================================
# CONFIG VALIDATION
# ===========================================================================

def _mask(value: str) -> str:
    value = value or ""
    return value if len(value) <= 8 else f"{value[:4]}...{value[-4:]}"

def validate_config() -> None:
    """Stop early with a clear message (in the log) instead of an HTTP error."""
    required = {}
    if RUN_ANAPLAN_EXPORT:
        required.update({"ANAPLAN_WORKSPACE_ID": WORKSPACE_ID, "ANAPLAN_MODEL_ID": MODEL_ID,
                         "ANAPLAN_EXPORT_ID": EXPORT_ID, "ANAPLAN_CLIENT_ID": ANAPLAN_CLIENT_ID})
    if PUSH_TO_FLEXERA:
        required.update({"FLEXERA_CLIENT_ID": CLIENT_ID,
                         "FLEXERA_CLIENT_SECRET": CLIENT_SECRET,
                         "FLEXERA_ASSIGNMENTS_CONFIG_ID": ASSIGNMENTS_CONFIG_ID,
                         "FLEXERA_ACTIVITIES_CONFIG_ID": ACTIVITIES_CONFIG_ID})
    missing = [k for k, v in required.items() if not v or v.startswith(("*", "PASTE_"))]
    if missing:
        sys.exit(f"Fill in these values in {ENV_FILE.name}: {', '.join(missing)}")
    if RUN_ANAPLAN_EXPORT and PUSH_TO_FLEXERA and ANAPLAN_CLIENT_ID == CLIENT_ID:
        sys.exit("ANAPLAN_CLIENT_ID and FLEXERA_CLIENT_ID in .env are identical. "
                 "ANAPLAN_CLIENT_ID must be the Anaplan OAuth client ID; "
                 "FLEXERA_CLIENT_ID must be the Flexera Application Registration ID.")
    if PUSH_TO_FLEXERA and ASSIGNMENTS_CONFIG_ID == ACTIVITIES_CONFIG_ID:
        sys.exit("FLEXERA_ASSIGNMENTS_CONFIG_ID and FLEXERA_ACTIVITIES_CONFIG_ID are identical "
                 "- each import needs its own config ID.")
    print(f"Settings: Anaplan export={'yes' if RUN_ANAPLAN_EXPORT else 'no'} | "
          f"Flexera push={'yes' if PUSH_TO_FLEXERA else 'no'} | "
          f"enabled users only={'yes' if ONLY_ENABLED_USERS else 'no'} | "
          f"zone normalise={'yes' if ZONE_NORMALISE else 'no'} | "
          f"blank zone -> {ZONE_BLANK_DEFAULT}"
          + (f" | unmapped zone -> {ZONE_UNMAPPED_DEFAULT}" if ZONE_NORMALISE else ""))
    if RUN_ANAPLAN_EXPORT:
        print(f"Anaplan client ID in use: {_mask(ANAPLAN_CLIENT_ID)}")
    if PUSH_TO_FLEXERA:
        print(f"Flexera client ID in use: {_mask(CLIENT_ID)} | "
              f"assignments config {_mask(ASSIGNMENTS_CONFIG_ID)} | "
              f"activities config {_mask(ACTIVITIES_CONFIG_ID)}")


# ===========================================================================
# ANAPLAN - HTTP helpers & authentication
# ===========================================================================

def raise_for_anaplan_error(response):
    """Raise a useful exception when an Anaplan request is unsuccessful."""
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        details = response.text.strip()
        if details:
            raise requests.HTTPError(
                f"{exc}. Response body: {details}", response=response
            ) from exc
        raise

def get_device_code(client_id, scope):
    """Request a device code for interactive user authentication."""
    response = requests.post(
        OAUTH_DEVICE_CODE_URL,
        headers={"Content-Type": "application/json"},
        json={"client_id": client_id, "scope": scope},
        timeout=60,
    )
    raise_for_anaplan_error(response)
    return response.json()

def authorize_user(verification_uri_complete):
    """Show the URL that the user must open to authorize the application."""
    print("\nPlease complete the Anaplan authorization process:")
    print(f"Open this URL in your browser: {verification_uri_complete}\n")

def get_anaplan_access_token(client_id, device_code, interval=5):
    """Poll Anaplan until device-code authorization is completed."""
    payload = {
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": device_code,
        "client_id": client_id,
    }
    while True:
        response = requests.post(
            OAUTH_TOKEN_URL,
            headers={"Content-Type": "application/json"},
            json=payload,
            timeout=60,
        )
        if response.status_code == 200:
            return response.json()
        if response.status_code in (400, 403):
            try:
                error = response.json().get("error")
            except (ValueError, json.JSONDecodeError):
                error = None
            if error == "authorization_pending":
                print("Waiting for user authorization...")
                time.sleep(interval)
                continue
            if error == "slow_down":
                interval += 5
                time.sleep(interval)
                continue
        raise_for_anaplan_error(response)

def refresh_access_token(client_id, refresh_token):
    """Get a new access token using a saved refresh token."""
    response = requests.post(
        OAUTH_TOKEN_URL,
        headers={"Content-Type": "application/json"},
        json={
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "client_id": client_id,
        },
        timeout=60,
    )
    raise_for_anaplan_error(response)
    tokens = response.json()
    tokens.setdefault("refresh_token", refresh_token)
    return tokens

def save_tokens(tokens, file_path=TOKEN_FILE):
    """Save OAuth tokens in the configured local directory."""
    file_path.parent.mkdir(parents=True, exist_ok=True)
    with file_path.open("w", encoding="utf-8") as token_file:
        json.dump(tokens, token_file, indent=2)

def load_tokens(file_path=TOKEN_FILE):
    """Load saved OAuth tokens, or return None when unavailable or invalid."""
    try:
        with file_path.open("r", encoding="utf-8") as token_file:
            content = token_file.read().strip()
        return json.loads(content) if content else None
    except (FileNotFoundError, json.JSONDecodeError):
        return None

def authenticate():
    """Return a valid Anaplan access token."""
    tokens = load_tokens()
    if tokens and tokens.get("refresh_token"):
        print("Using the saved refresh token...")
        try:
            tokens = refresh_access_token(ANAPLAN_CLIENT_ID, tokens["refresh_token"])
            save_tokens(tokens)
            print("Anaplan access token refreshed successfully.")
            return tokens["access_token"]
        except requests.RequestException as exc:
            print(f"Refresh failed: {exc}")
            print("Switching to device-code authorization...")

    try:
        device_response = get_device_code(ANAPLAN_CLIENT_ID, SCOPE)
    except requests.HTTPError as exc:
        if "unauthorized_client" in str(exc):
            sys.exit(
                "\nAnaplan rejected ANAPLAN_CLIENT_ID "
                f"({_mask(ANAPLAN_CLIENT_ID)}): 'Unauthorized or unknown client'.\n"
                "Fix: set ANAPLAN_CLIENT_ID to the Client ID of the Anaplan OAuth client\n"
                "(the value your original Anaplan export script used - NOT the Flexera\n"
                "Client ID). Check it in Anaplan Administration > Security > OAuth\n"
                "clients, and make sure the Device Grant flow is enabled for that client."
            )
        raise
    authorize_user(device_response["verification_uri_complete"])
    tokens = get_anaplan_access_token(
        ANAPLAN_CLIENT_ID,
        device_code=device_response["device_code"],
        interval=device_response.get("interval", 5),
    )
    save_tokens(tokens)
    print("Anaplan access token obtained successfully.")
    return tokens["access_token"]


# ===========================================================================
# ANAPLAN - CSV post-processing
# ===========================================================================

def parse_datetime_value(value):
    """Convert common Anaplan date/time strings to a datetime, or return None."""
    value = value.strip()
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        pass
    common_formats = (
        "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y",
        "%d %b %Y", "%d %B %Y", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
    )
    for date_format in common_formats:
        try:
            return datetime.strptime(value, date_format)
        except ValueError:
            continue
    return None

def parse_login_date(value):
    """Convert a supported date/time string to a date, or return None."""
    parsed = parse_datetime_value(value)
    return parsed.date() if parsed else None

def looks_like_date(value):
    """Return True when a non-empty value can be parsed as a date."""
    return bool(value.strip()) and parse_login_date(value) is not None

def post_process_export(output_file=OUTPUT_FILE):
    """Remove a date-only A1 row and add the 90-day inactivity column."""
    with output_file.open("r", encoding="utf-8-sig", newline="") as source:
        sample = source.read(4096)
        source.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;|\t")
        except csv.Error:
            dialect = csv.excel
        rows = list(csv.reader(source, dialect))

    if not rows:
        raise RuntimeError("The downloaded CSV is empty.")

    # A report period such as "Aug 26" can sit alone in A1 with the real
    # field names on row 2 - detect and drop it.
    first_row_nonempty = [cell.strip() for cell in rows[0] if cell.strip()]
    second_row_headers = (
        {cell.strip().casefold() for cell in rows[1] if cell.strip()}
        if len(rows) > 1 else set()
    )
    expected_second_row_headers = {
        "user id", "account creation date", "last active date", "license state",
    }
    first_row_is_metadata = (
        len(rows) > 1
        and len(first_row_nonempty) == 1
        and (
            looks_like_date(first_row_nonempty[0])
            or len(expected_second_row_headers & second_row_headers) >= 2
        )
    )
    if first_row_is_metadata:
        removed_value = first_row_nonempty[0]
        rows.pop(0)
        print(f"Removed the A1 metadata row: {removed_value!r}.")

    headers = [header.strip() for header in rows[0]]
    if not any(headers):
        raise RuntimeError("No CSV column headings were found after removing A1.")

    if not headers[0]:
        headers[0] = "Email ID"
        print('Added missing first-column header: "Email ID".')

    header_lookup = {header.casefold(): index for index, header in enumerate(headers)}
    login_index = header_lookup.get(LAST_LOGIN_COLUMN.casefold())
    if login_index is None:
        for candidate in ("last active date", "last login", "last login date",
                          "last login time", "last login datetime", "last logged in"):
            if candidate in header_lookup:
                login_index = header_lookup[candidate]
                break
    if login_index is None:
        raise ValueError(
            f'Could not find the last-login column "{LAST_LOGIN_COLUMN}". '
            f"Available columns: {', '.join(headers)}"
        )
    login_header = headers[login_index]

    # Drop previously calculated columns so reruns don't duplicate them.
    calculated_names = {INACTIVE_DAYS_COLUMN.casefold(), INACTIVE_COLUMN.casefold()}
    keep_indexes = [i for i, h in enumerate(headers) if h.casefold() not in calculated_names]
    if len(keep_indexes) != len(headers):
        headers = [headers[i] for i in keep_indexes]
        rows = [[row[i] if i < len(row) else "" for i in keep_indexes] for row in rows]
        login_index = headers.index(login_header)

    headers.extend([INACTIVE_DAYS_COLUMN, INACTIVE_COLUMN])
    inactive_days_index = len(headers) - 2
    inactive_index = len(headers) - 1
    current_headers = {header.casefold(): index for index, header in enumerate(headers)}

    today = date.today()
    processed_rows = [headers]
    for row_number, row in enumerate(rows[1:], start=2):
        row = list(row)
        if len(row) < len(headers):
            row.extend([""] * (len(headers) - len(row)))

        # Normalise date columns to ISO 8601 UTC, e.g. 2026-05-06T21:58:12Z
        for date_column in ISO_DATETIME_COLUMNS:
            date_index = current_headers.get(date_column.casefold())
            if date_index is None:
                raise ValueError(
                    f'Could not find required date column "{date_column}". '
                    f"Available columns: {', '.join(headers)}"
                )
            raw_date = row[date_index].strip() if date_index < len(row) else ""
            if raw_date:
                normalized = parse_datetime_value(raw_date)
                if normalized is None:
                    print(f"Warning: unrecognized {date_column} on CSV row "
                          f"{row_number}: {raw_date!r}")
                else:
                    if normalized.tzinfo is None:
                        normalized = normalized.replace(tzinfo=timezone.utc)
                    else:
                        normalized = normalized.astimezone(timezone.utc)
                    row[date_index] = normalized.isoformat(
                        timespec="seconds").replace("+00:00", "Z")

        raw_login = row[login_index] if login_index < len(row) else ""
        login_date = parse_login_date(raw_login)
        if not raw_login.strip():
            inactive_days, status = "", "Unknown"
        elif login_date is None:
            inactive_days, status = "", "Invalid date"
            print(f"Warning: unrecognized last-login date on CSV row {row_number}: {raw_login!r}")
        else:
            inactive_days = max(0, (today - login_date).days)
            status = "Yes" if inactive_days > INACTIVE_DAYS else "No"
        row[inactive_days_index] = inactive_days
        row[inactive_index] = status
        processed_rows.append(row)

    temporary_file = output_file.with_suffix(output_file.suffix + ".processed")
    with temporary_file.open("w", encoding="utf-8-sig", newline="") as destination:
        csv.writer(destination, dialect).writerows(processed_rows)
    temporary_file.replace(output_file)
    print(f'Added columns "{INACTIVE_DAYS_COLUMN}" and "{INACTIVE_COLUMN}" '
          f"using a {INACTIVE_DAYS}-day threshold.")


# ===========================================================================
# ANAPLAN - export & download
# ===========================================================================

def start_export(headers):
    """Start the configured Anaplan export and return its task ID."""
    url = (f"{ANAPLAN_API_BASE_URL}/workspaces/{WORKSPACE_ID}"
           f"/models/{MODEL_ID}/exports/{EXPORT_ID}/tasks")
    response = requests.post(url, headers=headers, json={"localeName": "en_US"}, timeout=60)
    raise_for_anaplan_error(response)
    task_id = response.json()["task"]["taskId"]
    print(f"Export started. Task ID: {task_id}")
    return task_id

def wait_for_export(headers, task_id, polling_interval=5):
    """Wait until the export task completes or fails."""
    url = (f"{ANAPLAN_API_BASE_URL}/workspaces/{WORKSPACE_ID}"
           f"/models/{MODEL_ID}/exports/{EXPORT_ID}/tasks/{task_id}")
    while True:
        response = requests.get(url, headers=headers, timeout=60)
        raise_for_anaplan_error(response)
        task = response.json()["task"]
        status = task["taskState"]
        print(f"Export status: {status}")
        if status == "COMPLETE":
            return
        if status in {"FAILED", "CANCELLED"}:
            raise RuntimeError(f"Anaplan export did not complete: {status}. {task.get('result', {})}")
        time.sleep(polling_interval)

def get_all_chunks(headers):
    """Retrieve the full list of chunks for the completed export file."""
    url = (f"{ANAPLAN_API_BASE_URL}/workspaces/{WORKSPACE_ID}"
           f"/models/{MODEL_ID}/files/{EXPORT_ID}/chunks")
    response = requests.get(url, headers=headers, timeout=60)
    raise_for_anaplan_error(response)
    return url, response.json().get("chunks", [])

def download_entire_dump(headers, output_file=OUTPUT_FILE):
    """Download every returned chunk and combine them into one output file."""
    chunks_url, chunks = get_all_chunks(headers)
    if not chunks:
        raise RuntimeError("The completed export returned no downloadable chunks.")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    temporary_file = output_file.with_suffix(output_file.suffix + ".part")
    try:
        with temporary_file.open("wb") as destination:
            total_chunks = len(chunks)
            for position, chunk in enumerate(chunks, start=1):
                chunk_id = str(chunk["id"])
                print(f"Downloading chunk {position}/{total_chunks}: {chunk_id}")
                chunk_headers = {
                    "Authorization": headers["Authorization"],
                    "Accept": "application/octet-stream",
                }
                with requests.get(f"{chunks_url}/{chunk_id}", headers=chunk_headers,
                                  timeout=300, stream=True) as response:
                    raise_for_anaplan_error(response)
                    for block in response.iter_content(chunk_size=1024 * 1024):
                        if block:
                            destination.write(block)
        temporary_file.replace(output_file)
    except Exception:
        temporary_file.unlink(missing_ok=True)
        raise
    print(f"Entire export dump downloaded to: {output_file}")
    print(f"Downloaded file size: {output_file.stat().st_size:,} bytes")

def run_anaplan_export():
    """Authenticate, run the export, download it and post-process it."""
    access_token = authenticate()
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    task_id = start_export(headers)
    wait_for_export(headers, task_id)
    download_entire_dump(headers)
    post_process_export(OUTPUT_FILE)


# ===========================================================================
# FLEXERA API (unchanged)
# ===========================================================================

def get_access_token() -> str:
    if not CLIENT_ID or not CLIENT_SECRET:
        sys.exit("FLEXERA_CLIENT_ID / FLEXERA_CLIENT_SECRET are not set.")
    resp = requests.post(
        TOKEN_URL,
        data={"grant_type": "client_credentials",
              "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET},
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    resp.raise_for_status()
    print("Obtained access token.")
    return resp.json()["access_token"]

def request_upload_url(token: str) -> tuple[str, str]:
    resp = requests.post(UPLOAD_URL,
                         headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    print("Received upload URL (fileId=%s, expires in %ss).", data.get("fileId"), data.get("expires"))
    return data["url"], data["fileId"]

def upload_csv_to_blob(presigned_url: str, csv_path: str) -> None:
    with open(csv_path, "rb") as f:
        resp = requests.put(
            presigned_url, data=f,
            headers={"x-ms-blob-type": "BlockBlob", "Content-Type": "text/csv"},
            timeout=120,
        )
    resp.raise_for_status()
    print("Uploaded CSV to blob storage (HTTP %s).", resp.status_code)

def trigger_import(token: str, file_id: str, config_id: str) -> None:
    IMPORT_URL = f"{BASE}/api/saas/import/v1/configs/{config_id}/records/import"
    resp = requests.post(
        IMPORT_URL,
        headers={"Authorization": f"Bearer {token}",
                 "Content-Type": "application/json"},
        data=json.dumps({"fileId": file_id}),
        timeout=60,
    )
    resp.raise_for_status()
    print("Import task created (HTTP %s). Response: %s", resp.status_code, resp.text[:500])

def make_user_id(email: str, org: str) -> str:
    key = f"{org.strip().lower()}\x00{email.strip().lower()}"
    return hashlib.md5(key.encode()).hexdigest()[:12]


# ===========================================================================
# BUILD FLEXERA FILES
# ===========================================================================

def _norm(s) -> str:
    return " ".join(str(s).strip().lower().split())

def load_export_for_flexera() -> pd.DataFrame:
    """Read the post-processed Anaplan export (delimiter auto-detected)."""
    if not OUTPUT_FILE.exists():
        sys.exit(f"Anaplan export not found: {OUTPUT_FILE}")
    df = pd.read_csv(OUTPUT_FILE, sep=None, engine="python", dtype=str,
                     keep_default_na=False, encoding="utf-8-sig")
    print(f"Loaded {len(df)} rows from {OUTPUT_FILE.name}.")

    rename, actual = {}, {_norm(c): c for c in df.columns}
    for expected, internal in SOURCE_COLUMNS.items():
        exp = _norm(expected)
        match = actual.get(exp) or next(
            (orig for a, orig in actual.items() if a.startswith(exp) and orig not in rename), None)
        if match is not None:
            rename[match] = internal
    df = df.rename(columns=rename)
    missing = [c for c in REQUIRED_SOURCE if c not in df.columns]
    if missing:
        sys.exit(f"Anaplan export is missing required column(s): {missing}. Found: {list(df.columns)}")
    return df

def _zone_key(value) -> str:
    """Uppercase, punctuation -> space, collapsed spaces."""
    cleaned = "".join(ch if ch.isalnum() else " " for ch in str(value).upper())
    return " ".join(cleaned.split())


def _load_extra_aliases() -> dict:
    """ZONE_EXTRA_ALIASES='Brasil=SAZ;Corporate Affairs=GHQ' from .env."""
    extra = {}
    for pair in env("ZONE_EXTRA_ALIASES").split(";"):
        if not pair.strip():
            continue
        if "=" not in pair:
            print(f"Warning: ignored ZONE_EXTRA_ALIASES entry {pair.strip()!r} (use variant=ZONE)")
            continue
        variant, zone = (part.strip() for part in pair.split("=", 1))
        zone = zone.upper()
        if zone not in CANONICAL_ZONES:
            print(f"Warning: ignored ZONE_EXTRA_ALIASES entry {pair.strip()!r} - "
                  f"zone must be one of {', '.join(CANONICAL_ZONES)}")
            continue
        extra[_zone_key(variant)] = zone
    return extra


def normalize_zone(value, aliases: dict, alias_order: list) -> tuple:
    """Return (canonical_zone, recognised: bool) for a non-blank zone."""
    key = _zone_key(value)
    padded = f" {key} "
    compact = key.replace(" ", "")
    # 1) whole-word alias match ("GHQ SOLUTIONS", "North America Zone")
    for alias in alias_order:
        if f" {alias} " in padded:
            return aliases[alias], True
    # 2) run-together spellings ("NorthAmerica", "SOUTHAMERICAZONE")
    for alias in alias_order:
        alias_compact = alias.replace(" ", "")
        if len(alias_compact) >= 4 and alias_compact in compact:
            return aliases[alias], True
    return None, False


def resolve_zones(raw: pd.DataFrame) -> pd.Series:
    """Zone per row, following ZONE_NORMALISE / ZONE_BLANK_DEFAULT /
    ZONE_UNMAPPED_DEFAULT in .env. Writes a summary to the log."""
    source = (raw["zone"] if "zone" in raw.columns
              else pd.Series("", index=raw.index)).fillna("").astype(str).str.strip()
    if "zone" not in raw.columns:
        print(f"Warning: no 'Assigned Zone' column in the export - every user gets "
              f"ZONE_BLANK_DEFAULT ({ZONE_BLANK_DEFAULT})")
    is_blank = source.map(lambda v: _zone_key(v) in BLANK_ZONE_VALUES)

    if not ZONE_NORMALISE:
        zones = source.where(~is_blank, ZONE_BLANK_DEFAULT)
        print(f"Zone normalisation OFF - Assigned Zone kept as-is; "
              f"{int(is_blank.sum())} blank -> {ZONE_BLANK_DEFAULT}")
        counts = zones.value_counts()
        print("Zones: " + ", ".join(f"{z}={n}" for z, n in counts.items()))
        return zones

    aliases = {**ZONE_ALIASES, **_load_extra_aliases()}
    alias_order = sorted(aliases, key=lambda a: (-len(a.split()), -len(a)))
    keep_unmapped = ZONE_UNMAPPED_DEFAULT.upper() == "KEEP"

    zones, unmapped = [], []
    for value, blank in zip(source, is_blank):
        if blank:
            zones.append(ZONE_BLANK_DEFAULT)
            continue
        zone, recognised = normalize_zone(value, aliases, alias_order)
        if not recognised:
            unmapped.append(value)
            zone = value if keep_unmapped else ZONE_UNMAPPED_DEFAULT
        zones.append(zone)
    zones = pd.Series(zones, index=raw.index)

    counts = zones.value_counts()
    others = {z: n for z, n in counts.items() if z not in CANONICAL_ZONES}
    summary = ", ".join(f"{z}={counts.get(z, 0)}" for z in CANONICAL_ZONES)
    if others:
        summary += ", other: " + ", ".join(f"{z}={n}" for z, n in others.items())
    print(f"Zone normalisation ON - {summary} "
          f"({int(is_blank.sum())} blank -> {ZONE_BLANK_DEFAULT})")
    if unmapped:
        target = "kept as-is" if keep_unmapped else f"set to {ZONE_UNMAPPED_DEFAULT}"
        print(f"Warning: {len(unmapped)} row(s) with an unrecognised Assigned Zone {target} - "
              f"add them to ZONE_EXTRA_ALIASES in .env: {sorted(set(unmapped))}")
    return zones


def build_flexera_frames(raw: pd.DataFrame):
    """Return (assignments, activities) with Flexera headers, dates untouched."""
    raw = raw.copy()
    raw["email"] = raw["email"].str.strip().str.lower()
    raw = raw[raw["email"] != ""]

    if ONLY_ENABLED_USERS and "user_state" in raw.columns:
        before = len(raw)
        raw = raw[raw["user_state"].str.strip().str.lower() == "enabled"]
        print(f"ONLY_ENABLED_USERS: kept {len(raw)} of {before} users")

    ids = raw["anaplan_id"].str.strip() if "anaplan_id" in raw.columns else pd.Series("", index=raw.index)
    raw["user_id"] = [aid if aid else make_user_id(email, ANAPLAN_TENANT_KEY)
                      for aid, email in zip(ids, raw["email"])]
    generated = int((ids == "").sum())
    if generated:
        print(f"Warning: {generated} user(s) had no Anaplan User ID - generated with make_user_id()")

    raw["zone_final"] = resolve_zones(raw)
    raw["license_final"] = [
        LICENSE_ZONE_SEPARATOR.join(p for p in (lt.strip(), zone.strip()) if p)
        for lt, zone in zip(raw["license_type"], raw["zone_final"])
    ]
    created = raw["created"].str.strip() if "created" in raw.columns else ""
    last_active = raw["last_active"].str.strip()

    assignments = pd.DataFrame({
        "Email ID":              raw["email"],
        "User ID":               raw["user_id"],
        "License Type":          raw["license_final"],
        "Account Creation Date": created,
        "Last Active Date":      last_active,
    })[ASSIGNMENTS_TEMPLATE]
    assignments = assignments.drop_duplicates(subset=["Email ID", "License Type"]).reset_index(drop=True)

    activities = pd.DataFrame({"User ID": raw["user_id"], "Last Active Date": last_active})
    no_activity = int((activities["Last Active Date"] == "").sum())
    activities = activities[activities["Last Active Date"] != ""]
    activities = (activities
                  .assign(_sort=pd.to_datetime(activities["Last Active Date"], errors="coerce", utc=True))
                  .sort_values("_sort", ascending=False, na_position="last")
                  .drop_duplicates(subset=["User ID"], keep="first")
                  .drop(columns="_sort"))[ACTIVITIES_TEMPLATE].reset_index(drop=True)
    if no_activity:
        print(f"{no_activity} user(s) with no Last Active Date left out of activities")

    license_counts = assignments["License Type"].value_counts()
    print("License Types: " + ", ".join(f"{lt}={n}" for lt, n in license_counts.items()))
    return assignments, activities


def push_to_flexera(csv_path: Path, config_id: str, label: str) -> None:
    print(f"--- Pushing {label} ({csv_path.name}) ---")
    access_token = get_access_token()
    presigned_url, file_id = request_upload_url(access_token)
    upload_csv_to_blob(presigned_url, str(csv_path))
    trigger_import(access_token, file_id, config_id=config_id)


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    started = time.time()
    setup_logging()
    print("=" * 70)
    print(f"Anaplan -> Flexera sync started | .env: {ENV_FILE} | log: {LOG_PATH}")
    BASE_DIRECTORY.mkdir(parents=True, exist_ok=True)
    validate_config()

    if RUN_ANAPLAN_EXPORT:
        run_anaplan_export()
    else:
        print(f"RUN_ANAPLAN_EXPORT=no - using existing {OUTPUT_FILE.name}.")

    assignments, activities = build_flexera_frames(load_export_for_flexera())

    # Flexera CSVs: fixed names, plain UTF-8, comma-delimited
    assignments.to_csv(ASSIGNMENTS_CSV, index=False, encoding="utf-8")
    activities.to_csv(ACTIVITIES_CSV, index=False, encoding="utf-8")
    print(f"Wrote {len(assignments)} assignment rows -> {ASSIGNMENTS_CSV.name} {list(assignments.columns)}")
    print(f"Wrote {len(activities)} activity rows   -> {ACTIVITIES_CSV.name} {list(activities.columns)}")

    if not PUSH_TO_FLEXERA:
        print("FLEXERA_PUSH=no - files built, nothing sent.")
    else:
        if len(assignments):
            push_to_flexera(ASSIGNMENTS_CSV, ASSIGNMENTS_CONFIG_ID, "user assignments")
        if len(activities):
            push_to_flexera(ACTIVITIES_CSV, ACTIVITIES_CONFIG_ID, "user activities")
        print("Pushed. Check Flexera > Data imports > Anaplan for the import result.")
    print(f"Finished OK in {time.time() - started:.1f}s")


def run():
    """Run main() so every outcome - success, stop or crash - lands in the log."""
    try:
        main()
    except SystemExit as exc:
        if exc.code not in (None, 0):
            setup_logging()
            LOG.error(f"Run stopped: {exc.code}")
            sys.exit(1)
    except KeyboardInterrupt:
        setup_logging()
        LOG.error("Run cancelled by user (Ctrl+C)")
        sys.exit(1)
    except Exception:
        setup_logging()
        LOG.exception("Run FAILED with an unexpected error")
        sys.exit(1)


if __name__ == "__main__":
    run()
