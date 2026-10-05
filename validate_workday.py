"""
Validation checks for the Workday headcount / termination views.
Answers: are there rehires, what status values exist, how old are the leavers,
and WHY there are duplicate rows (which columns differ between them).

Reads the same .env as get_workday_leavers.py. Run it the same way:
    python validate_workday.py

Prints each check and saves them to output/validation_<date>.xlsx
"""
import os
import sys
from datetime import date

import pandas as pd
from databricks import sql


def load_env(path=None):
    path = path or os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, _, v = line.partition("=")
                    k, v = k.strip(), v.strip().strip('"').strip("'")
                    if k and k not in os.environ:
                        os.environ[k] = v


load_env()
env = lambda k, d="": os.environ.get(k, d).strip()

HOST = env("DATABRICKS_HOST", "adb-3060038306817150.10.azuredatabricks.net")
HTTP_PATH = env("DATABRICKS_HTTP_PATH", "/sql/1.0/warehouses/13861ac29ff00f3e")
CATALOG = env("DATABRICKS_CATALOG", "brewdat_uc_people_prod")
SCHEMA = env("DATABRICKS_SCHEMA", "gld_ghq_people_workday_tr")
TOKEN = env("DATABRICKS_TOKEN")
OUTPUT_DIR = env("OUTPUT_DIR", "output")

HC = f"{CATALOG}.{SCHEMA}.people_psi_headcount_sam_coe_vw"
TM = f"{CATALOG}.{SCHEMA}.people_psi_termination_sam_coe_vw"

CHECKS = {
    # 1. Which employee_status values exist, and how many of each
    "status_values": f"""
        SELECT employee_status, COUNT(*) AS rows,
               COUNT(DISTINCT global_employee_id) AS employees
        FROM {HC}
        GROUP BY employee_status
        ORDER BY rows DESC
    """,
    # 2. People in BOTH views - the possible rehires
    "in_both_views": f"""
        SELECT h.employee_status, COUNT(DISTINCT h.global_employee_id) AS employees
        FROM {HC} h
        JOIN {TM} t ON h.global_employee_id = t.global_employee_id
        GROUP BY h.employee_status
        ORDER BY employees DESC
    """,
    # 3. Rehires: still Active today but with a termination date in the past
    "active_with_past_termination": f"""
        SELECT COUNT(DISTINCT h.global_employee_id) AS employees
        FROM {HC} h
        JOIN {TM} t ON h.global_employee_id = t.global_employee_id
        WHERE lower(h.employee_status) = 'active'
          AND to_date(t.termination_date) < current_date()
    """,
    # 4. How old the leavers are - helps choose a cutoff date
    "leavers_by_year": f"""
        SELECT year(to_date(termination_date)) AS termination_year,
               COUNT(DISTINCT global_employee_id) AS employees
        FROM {TM}
        GROUP BY year(to_date(termination_date))
        ORDER BY termination_year DESC
    """,
    # 5. Row counts vs unique employees (the duplicate picture)
    "row_vs_unique": f"""
        SELECT 'headcount' AS view_name, COUNT(*) AS rows,
               COUNT(DISTINCT global_employee_id) AS employees
        FROM {HC}
        UNION ALL
        SELECT 'termination', COUNT(*), COUNT(DISTINCT global_employee_id)
        FROM {TM}
    """,
}


def duplicate_columns(conn, view, label, writer):
    """For people with more than one row, which columns actually differ?"""
    query = f"""
        SELECT * FROM {view}
        WHERE global_employee_id IN (
            SELECT global_employee_id FROM {view}
            WHERE global_employee_id IS NOT NULL
            GROUP BY global_employee_id HAVING COUNT(*) > 1)
    """
    with conn.cursor() as cur:
        cur.execute(query)
        dup = cur.fetchall_arrow().to_pandas()

    people = dup["global_employee_id"].nunique()
    print(f"\n=== duplicates_{label} ===")
    print(f"{len(dup)} rows belong to {people} people with more than one row")
    if not people:
        return

    # Number of people whose rows disagree on each column
    differs = (dup.groupby("global_employee_id").nunique(dropna=False) > 1).sum()
    summary = (differs[differs > 0].sort_values(ascending=False)
               .rename("people_with_different_values").to_frame())
    summary["pct_of_duplicated_people"] = (
        summary["people_with_different_values"] / people * 100).round(1)
    summary.index.name = "column"
    summary = summary.reset_index()
    print(summary.to_string(index=False) if len(summary)
          else "(rows are exact copies - every column identical)")

    # Rows per person, to see whether it is 2 rows each or a few with many
    spread = (dup.groupby("global_employee_id").size().value_counts()
              .sort_index().rename("people").to_frame())
    spread.index.name = "rows_per_person"
    print(spread.reset_index().to_string(index=False))

    summary.to_excel(writer, sheet_name=f"dup_columns_{label}"[:31], index=False)
    spread.reset_index().to_excel(writer, sheet_name=f"dup_spread_{label}"[:31],
                                  index=False)
    # A handful of real examples to eyeball (first 15 people)
    sample_ids = dup["global_employee_id"].drop_duplicates().head(15)
    (dup[dup["global_employee_id"].isin(sample_ids)]
     .sort_values("global_employee_id")
     .to_excel(writer, sheet_name=f"dup_sample_{label}"[:31], index=False))


def main():
    if not TOKEN:
        sys.exit("No DATABRICKS_TOKEN found - check your .env file")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out = os.path.join(OUTPUT_DIR, f"validation_{date.today():%Y%m%d}.xlsx")

    with sql.connect(server_hostname=HOST, http_path=HTTP_PATH,
                     access_token=TOKEN) as conn, \
            pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, query in CHECKS.items():
            with conn.cursor() as cur:
                cur.execute(query)
                df = cur.fetchall_arrow().to_pandas()
            print(f"\n=== {name} ===")
            print(df.to_string(index=False) if len(df) else "(no rows)")
            df.to_excel(writer, sheet_name=name[:31], index=False)

        # 6 & 7. What makes the duplicate rows different
        duplicate_columns(conn, HC, "headcount", writer)
        duplicate_columns(conn, TM, "termination", writer)

    print(f"\nSaved to {out}")


if __name__ == "__main__":
    main()
