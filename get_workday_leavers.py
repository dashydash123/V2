"""
Scheduled pull of the Workday SAM CoE views from the Databricks SQL warehouse.

Every table/view in the schema goes into its own sheet of one Excel workbook,
plus an HR_Roster sheet (Flexera HR roster format) built from the headcount and
termination views, an HR_Roster_Mapping sheet, and a Flexera-ready CSV.

Setup:
    pip install databricks-sql-connector databricks-sdk pandas pyarrow openpyxl xlsxwriter
    (xlsxwriter is optional but uses much less memory for the Excel file)

All settings live in a .env file next to this script (see .env.example).
Nothing secret is written to the log.
"""
import os
import re
import sys
import logging
from datetime import date, datetime
from logging.handlers import RotatingFileHandler

import pandas as pd
from databricks import sql
from databricks.sdk.core import Config, oauth_service_principal


# ---------------- .env loading ----------------
def load_env(path=None):
    """Read KEY=VALUE lines from .env. Real environment variables win."""
    path = path or os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(path):
        return path, False
    with open(path, encoding="utf-8-sig") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key, value = key.strip(), value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value
    return path, True


ENV_PATH, ENV_FOUND = load_env()


def env(key, default=""):
    return os.environ.get(key, default).strip()


# ---------------- settings ----------------
HOST = env("DATABRICKS_HOST", "adb-3060038306817150.10.azuredatabricks.net")
HTTP_PATH = env("DATABRICKS_HTTP_PATH", "/sql/1.0/warehouses/13861ac29ff00f3e")
CATALOG = env("DATABRICKS_CATALOG", "brewdat_uc_people_prod")
SCHEMA = env("DATABRICKS_SCHEMA", "gld_ghq_people_workday_tr")

TOKEN = env("DATABRICKS_TOKEN")
CLIENT_ID = env("DATABRICKS_CLIENT_ID")
CLIENT_SECRET = env("DATABRICKS_CLIENT_SECRET")

# Leave blank for no cutoff. yyyy-mm-dd: leavers whose latest termination date
# is BEFORE this date are left out of the HR roster (raw sheets keep everything).
CUTOFF_DATE = env("TERMINATION_CUTOFF_DATE")

# Rehire rule: someone present in the headcount view (Active / OnLeave) is
# treated as employed even if the termination view holds a past date for them.
# Grace period guards against snapshot lag: a termination this many days old
# or newer still counts as a leaver. 0 = headcount always wins.
REHIRE_GRACE_DAYS = int(env("REHIRE_GRACE_DAYS", "0") or 0)

OUTPUT_DIR = env("OUTPUT_DIR", "output")
LOG_FILE = env("LOG_FILE", os.path.join(OUTPUT_DIR, "workday_hr_roster.log"))
LOG_MAX_BYTES = 5 * 1024 * 1024

EXCEL_MAX_ROWS = 1_048_575  # Excel limit minus header row
ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# ---------------- HR roster configuration ----------------
HEADCOUNT_VIEW = "people_psi_headcount_sam_coe_vw"
TERMINATION_VIEW = "people_psi_termination_sam_coe_vw"

ROSTER_COLUMNS = ["id", "email", "employmentStatus", "name", "country",
                  "location", "department", "employeeId", "employeeType",
                  "hireDate", "terminationDate", "costCenter", "isDeleted"]

# Output field -> source columns, in priority order (first non-blank wins).
FIELD_SOURCES = {
    "email":        ["employee_email", "upn"],
    "name":         ["employee_name"],
    "country":      ["country_name", "country"],
    "location":     ["zone"],
    "department":   ["function_l1_desc"],
    "employeeId":   ["local_employee_id"],
    "employeeType": [],                     # left blank for now
    "hireDate":     [],                     # left blank for now
    "costCenter":   ["cost_center_desc", "cost_center"],
}
JOIN_KEY = "global_employee_id"           # also used as the roster "id"
HC_STATUS_COL = "employee_status"
HC_LOAD_DATE_COL = "load_date"            # used to pick the latest duplicate
TERM_DATE_COL = "termination_date"

MAPPING_NOTES = {
    "id": "global_employee_id (join key). Rows with no global ID are excluded.",
    "email": "employee_email, falls back to upn when blank.",
    "employmentStatus": "Still in the headcount view -> employee_status wins "
                        "(Active -> Employed, on leave -> On leave), even with a "
                        "past termination date (rehire). Otherwise in termination "
                        "view -> Not employed; future termination date -> Employed; "
                        "no headcount row and no status -> Not monitored.",
    "location": "zone (e.g. ZONE EUROPE).",
    "department": "function_l1_desc - CONFIRM this is the department level wanted.",
    "employeeId": "local_employee_id - CONFIRM (same as global ID for many rows).",
    "employeeType": "Left blank for now (no agreed value mapping yet).",
    "hireDate": "Left blank for now (no hire date column in the views).",
    "terminationDate": "termination_date (latest per employee), yyyy-mm-dd; "
                       "future dates are kept; cleared for rehires.",
    "costCenter": "cost_center_desc - this is a description, not a code.",
    "isDeleted": "TRUE when employmentStatus is Not employed, else FALSE.",
}

log = logging.getLogger("workday")


# ---------------- logging ----------------
def setup_logging():
    os.makedirs(os.path.dirname(os.path.abspath(LOG_FILE)) or ".", exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s",
                            "%Y-%m-%d %H:%M:%S")
    fh = RotatingFileHandler(LOG_FILE, maxBytes=LOG_MAX_BYTES, backupCount=1,
                             encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(sh)

    log.info("=" * 78)
    log.info("Run started %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log.info("Settings file: %s", ENV_PATH if ENV_FOUND else f"{ENV_PATH} (not found - using defaults)")
    log.info("Warehouse: %s%s", HOST, HTTP_PATH)
    log.info("Schema: %s.%s", CATALOG, SCHEMA)
    log.info("Termination cutoff: %s", CUTOFF_DATE or "none (all leavers)")
    log.info("Rehire rule: headcount wins%s",
             f" (grace {REHIRE_GRACE_DAYS} days)" if REHIRE_GRACE_DAYS else "")


# ---------------- connection ----------------
def connect():
    if CLIENT_ID and CLIENT_SECRET:
        log.info("Auth: service principal (client ID %s...)", CLIENT_ID[:8])

        def credentials_provider():
            cfg = Config(host=f"https://{HOST}", client_id=CLIENT_ID,
                         client_secret=CLIENT_SECRET)
            return oauth_service_principal(cfg)
        return sql.connect(server_hostname=HOST, http_path=HTTP_PATH,
                           credentials_provider=credentials_provider)
    if TOKEN:
        log.info("Auth: personal access token")
        return sql.connect(server_hostname=HOST, http_path=HTTP_PATH,
                           access_token=TOKEN)
    raise ValueError("No credentials: set DATABRICKS_TOKEN (or "
                     "DATABRICKS_CLIENT_ID/SECRET) in .env")


def list_objects(cur):
    """All tables and views in the schema that this identity can read."""
    cur.execute(f"""
        SELECT table_name, table_type
        FROM {CATALOG}.information_schema.tables
        WHERE table_schema = '{SCHEMA}'
        ORDER BY table_name
    """)
    return cur.fetchall()


def sheet_name(name, used):
    """Excel sheet names: max 31 chars, no []:*?/\\ , must be unique."""
    base = re.sub(r"[\[\]:*?/\\]", "_", name)[:31]
    candidate, n = base, 1
    while candidate.lower() in used:
        suffix = f"~{n}"
        candidate = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def clean_for_excel(df):
    """Strip timezones and control characters that Excel/openpyxl reject."""
    for col in df.columns:
        if isinstance(df[col].dtype, pd.DatetimeTZDtype):
            df[col] = df[col].dt.tz_localize(None)
        elif pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].map(
                lambda v: ILLEGAL_CHARS.sub("", v) if isinstance(v, str) else v)
    return df


# ---------------- HR roster build ----------------
def _lower_cols(df):
    return df.rename(columns=lambda c: str(c).strip().lower())


def _text(series):
    s = series.astype("string").str.strip()
    return s.mask(s == "", pd.NA)


def _id(series):
    return _text(series).str.replace(r"\.0$", "", regex=True)


def _to_date(series):
    return pd.to_datetime(series, errors="coerce", format="mixed")


def _coalesce(df, candidates):
    """First non-blank value across candidate columns; returns (series, used)."""
    out = pd.Series(pd.NA, index=df.index, dtype="string")
    used = []
    for col in candidates:
        if col in df.columns:
            used.append(col)
            out = out.fillna(_text(df[col]))
    return out, used


def _headcount_status(value):
    if pd.isna(value):
        return "Not monitored"
    key = re.sub(r"[\s_\-]", "", str(value)).lower()
    if key == "active":
        return "Employed"
    if "leave" in key:
        return "On leave"
    return "Not monitored"


def build_hr_roster(headcount, termination, cutoff=None):
    """Returns (roster_df, mapping_df, stats_df)."""
    today = pd.Timestamp(date.today())
    hc = _lower_cols(headcount)
    tm = _lower_cols(termination)
    for name, df in (("headcount", hc), ("termination", tm)):
        if JOIN_KEY not in df.columns:
            raise ValueError(f"{JOIN_KEY} missing from {name} view")

    # Headcount: drop rows without an ID, keep latest duplicate by load_date
    hc["_gid"] = _id(hc[JOIN_KEY])
    hc_no_id = int(hc["_gid"].isna().sum())
    hc = hc[hc["_gid"].notna()]
    if HC_LOAD_DATE_COL in hc.columns:
        hc = hc.assign(_ld=_to_date(hc[HC_LOAD_DATE_COL])).sort_values(
            "_ld", ascending=False, na_position="last")
    hc_before = len(hc)
    hc = hc.drop_duplicates("_gid", keep="first")
    hc_dupes = hc_before - len(hc)
    log.info("Headcount: %s rows -> %s unique employees (%s no ID, %s duplicates)",
             len(headcount), len(hc), hc_no_id, hc_dupes)

    # Termination: keep the latest termination date per employee
    tm["_gid"] = _id(tm[JOIN_KEY])
    tm_no_id = int(tm["_gid"].isna().sum())
    tm = tm[tm["_gid"].notna()]
    tm["_tdate"] = _to_date(tm[TERM_DATE_COL]) if TERM_DATE_COL in tm.columns \
        else pd.NaT
    tm = tm.sort_values("_tdate", ascending=False, na_position="last")
    tm_before = len(tm)
    tm = tm.drop_duplicates("_gid", keep="first")
    tm_dupes = tm_before - len(tm)
    log.info("Termination: %s rows -> %s unique leavers (%s no ID, %s duplicates)",
             len(termination), len(tm), tm_no_id, tm_dupes)

    # Optional cutoff: ignore leavers who left before it
    cut_excluded = 0
    if cutoff is not None:
        keep = tm["_tdate"].isna() | (tm["_tdate"] >= cutoff)
        cut_excluded = int((~keep).sum())
        tm = tm[keep]
        log.info("Cutoff %s: %s older leavers left out, %s kept",
                 cutoff.date(), cut_excluded, len(tm))

    def project(df, source):
        out = pd.DataFrame(index=df.index)
        out["id"] = df["_gid"]
        used = {}
        for field, cands in FIELD_SOURCES.items():
            out[field], used[field] = _coalesce(df, cands)
        out["_hc_status"] = df[HC_STATUS_COL] if (
            source == "hc" and HC_STATUS_COL in df.columns) else pd.NA
        return out, used

    hc_out, hc_used = project(hc, "hc")
    tm_only = tm[~tm["_gid"].isin(hc["_gid"])]
    tm_out, tm_used = project(tm_only, "tm")

    roster = pd.concat([hc_out, tm_out], ignore_index=True)
    tdates = tm.set_index("_gid")["_tdate"]
    roster["_tdate"] = roster["id"].map(tdates)
    in_term = roster["id"].isin(tm["_gid"])

    in_hc = pd.Series([True] * len(hc_out) + [False] * len(tm_out),
                      index=roster.index)
    grace = pd.Timedelta(days=REHIRE_GRACE_DAYS)

    def status(row, termed, in_headcount):
        tdate = row["_tdate"]
        if termed and pd.notna(tdate) and tdate > today:
            return "Employed"              # future leaver - still employed
        if in_headcount:
            # Rehire rule: a past termination is overridden by the live
            # headcount row, unless it is inside the grace period.
            if termed and REHIRE_GRACE_DAYS and pd.notna(tdate) \
                    and tdate >= today - grace:
                return "Not employed"
            return _headcount_status(row["_hc_status"])
        if termed:
            return "Not employed"
        return _headcount_status(row["_hc_status"])

    roster["employmentStatus"] = [
        status(r, t, h) for (_, r), t, h
        in zip(roster.iterrows(), in_term, in_hc)]
    past_term = in_term & (roster["_tdate"] <= today)
    future_dated = int((in_term & (roster["_tdate"] > today)).sum())
    rehires = int((past_term & in_hc &
                   (roster["employmentStatus"] != "Not employed")).sum())
    log.info("Rehires kept employed (past termination, still in headcount): %s",
             rehires)
    # A rehire's old termination date would be misleading next to an active
    # status, so only keep dates for leavers and future-dated terminations.
    keep_tdate = (roster["employmentStatus"] == "Not employed") | (
        roster["_tdate"] > today)
    roster["terminationDate"] = roster["_tdate"].where(
        keep_tdate).dt.strftime("%Y-%m-%d")
    roster["hireDate"] = _to_date(roster["hireDate"]).dt.strftime("%Y-%m-%d")
    roster["isDeleted"] = (roster["employmentStatus"] == "Not employed").map(
        {True: "TRUE", False: "FALSE"})
    roster = roster[ROSTER_COLUMNS]

    mapping = []
    for field in ROSTER_COLUMNS:
        hc_src = ", ".join(hc_used.get(field, [])) or (
            JOIN_KEY if field == "id" else "")
        tm_src = ", ".join(tm_used.get(field, [])) or (
            JOIN_KEY if field == "id" else "")
        if field in ("employmentStatus", "isDeleted"):
            hc_src, tm_src = HC_STATUS_COL, TERM_DATE_COL
        if field == "terminationDate":
            hc_src, tm_src = "", TERM_DATE_COL
        mapped = "Yes" if (hc_src or tm_src) else "Blank"
        mapping.append([field, hc_src, tm_src, mapped,
                        MAPPING_NOTES.get(field, "")])
    mapping_df = pd.DataFrame(mapping, columns=[
        "Roster field", "Headcount source", "Termination source",
        "Mapped", "Notes"])

    counts = roster["employmentStatus"].value_counts()
    stats_df = pd.DataFrame([
        ["Total roster rows", len(roster)],
        ["Employed", int(counts.get("Employed", 0))],
        ["On leave", int(counts.get("On leave", 0))],
        ["Not employed", int(counts.get("Not employed", 0))],
        ["Not monitored", int(counts.get("Not monitored", 0))],
        ["Leavers only in termination view (added)", len(tm_out)],
        ["Rehires (past termination, kept employed)", rehires],
        ["Rehire grace period (days)", REHIRE_GRACE_DAYS],
        ["Future-dated leavers (still Employed)", future_dated],
        ["Headcount duplicates removed", hc_dupes],
        ["Termination duplicates removed", tm_dupes],
        ["Headcount rows excluded (no global ID)", hc_no_id],
        ["Termination rows excluded (no global ID)", tm_no_id],
        ["Termination cutoff date", CUTOFF_DATE or "none"],
        ["Leavers excluded by cutoff", cut_excluded],
        ["Run date", date.today().isoformat()],
    ], columns=["Check", "Value"])
    for _, r in stats_df.iterrows():
        log.info("  %-42s %s", r["Check"], r["Value"])
    return roster, mapping_df, stats_df


def excel_engine():
    """xlsxwriter uses far less memory than openpyxl when it's installed."""
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except ImportError:
        return "openpyxl"


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out = os.path.join(OUTPUT_DIR,
                       f"workday_{SCHEMA}_{date.today():%Y%m%d}.xlsx")
    csv_out = os.path.join(OUTPUT_DIR, f"hr_roster_{date.today():%Y%m%d}.csv")
    cutoff = pd.Timestamp(CUTOFF_DATE) if CUTOFF_DATE else None

    index_rows, failures = [], 0
    used = {"index", "hr_roster", "hr_roster_mapping"}
    frames, plan = {}, []          # plan: (sheet names, dataframe)

    # 1. Fetch every view first, then close the connection
    with connect() as conn, conn.cursor() as cur:
        objects = list_objects(cur)
        if not objects:
            raise ValueError(f"No readable tables/views found in {CATALOG}.{SCHEMA}")
        log.info("Found %s objects in the schema", len(objects))

        for table_name, table_type in objects:
            full = f"{CATALOG}.{SCHEMA}.{table_name}"
            try:
                cur.execute(f"SELECT * FROM {full}")
                df = clean_for_excel(cur.fetchall_arrow().to_pandas())
                frames[table_name.lower()] = df
                chunks = range(0, max(len(df), 1), EXCEL_MAX_ROWS)
                sheets = [sheet_name(table_name if i == 0 else
                                     f"{table_name}_{i + 1}", used)
                          for i, _ in enumerate(chunks)]
                plan.append((sheets, df))
                index_rows.append([", ".join(sheets), full, table_type,
                                   len(df), "OK"])
                log.info("%s: %s rows", full, len(df))
            except Exception as e:
                failures += 1
                index_rows.append(["", full, table_type, None, f"FAILED: {e}"])
                log.error("%s: FAILED - %s", full, e)

    # 2. Build the Flexera HR roster and its CSV
    roster = mapping_df = stats_df = None
    try:
        if HEADCOUNT_VIEW not in frames or TERMINATION_VIEW not in frames:
            raise ValueError("headcount or termination view not loaded")
        roster, mapping_df, stats_df = build_hr_roster(
            frames[HEADCOUNT_VIEW], frames[TERMINATION_VIEW], cutoff)
        roster.to_csv(csv_out, index=False, encoding="utf-8-sig")
        index_rows.insert(0, ["HR_Roster, HR_Roster_Mapping",
                              f"Built from {HEADCOUNT_VIEW} + {TERMINATION_VIEW}",
                              "DERIVED", len(roster),
                              f"OK - CSV: {os.path.basename(csv_out)}"])
        log.info("HR_Roster: %s rows -> %s", len(roster), csv_out)
    except Exception as e:
        failures += 1
        index_rows.insert(0, ["HR_Roster", "Derived roster", "DERIVED",
                              None, f"FAILED: {e}"])
        log.error("HR_Roster: FAILED - %s", e)

    # 3. Write the workbook once, already in the right sheet order:
    #    Index, HR_Roster, HR_Roster_Mapping, then the raw views.
    #    (No re-opening afterwards - that is what ran out of memory.)
    engine = excel_engine()
    log.info("Writing %s (engine: %s)...", out, engine)
    with pd.ExcelWriter(out, engine=engine) as writer:
        pd.DataFrame(index_rows, columns=["Sheet", "Source object", "Type",
                                          "Rows", "Status"]
                     ).to_excel(writer, sheet_name="Index", index=False)
        if roster is not None:
            roster.to_excel(writer, sheet_name="HR_Roster", index=False)
            mapping_df.to_excel(writer, sheet_name="HR_Roster_Mapping",
                                index=False)
            stats_df.to_excel(writer, sheet_name="HR_Roster_Mapping",
                              index=False, startrow=len(mapping_df) + 3)
        for sheets, df in plan:
            for i, name in enumerate(sheets):
                start = i * EXCEL_MAX_ROWS
                df.iloc[start:start + EXCEL_MAX_ROWS].to_excel(
                    writer, sheet_name=name, index=False)

    log.info("Saved %s objects to %s", len(objects), out)
    if failures:
        raise RuntimeError(f"{failures} object(s) failed - see Index sheet")
    log.info("Run finished OK")


if __name__ == "__main__":
    setup_logging()
    try:
        main()
    except Exception as e:
        log.exception("Run FAILED: %s", e)
        sys.exit(1)
