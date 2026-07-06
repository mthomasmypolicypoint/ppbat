"""
Benefit Activation Read Out Generator
--------------------------------------
- Reads: Detail Report.xlsx, PolicyIndividualsGroups (48).xlsx, Calls.xlsx
- Outputs: Benefit Activation Read Out.xlsx  (one level up from this script)
- Versions prior output as: BAT MM.DD.YYYY.xlsx (date = file's last-modified date)
"""

import os
import re
import shutil
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Paths ──────────────────────────────────────────────────────────────────────
DATA_DIR = r"C:\Users\miche\OneDrive - My Policy Point\PolicyPoint Leadership Library - Documents\Member Experience\BAT"

DETAIL_FILE = os.path.join(DATA_DIR, "Detail Report.csv")
POLICY_FILE = os.path.join(DATA_DIR, "PolicyIndividualsGroups.csv")
CALLS_FILE  = os.path.join(DATA_DIR, "Calls.csv")
OUTPUT_FILE = os.path.join(DATA_DIR, "Benefit Activation Read Out.xlsx")


# ── Helpers ────────────────────────────────────────────────────────────────────
def normalize(value):
    """Strip all non-alphanumeric characters and uppercase."""
    if pd.isna(value):
        return ""
    return re.sub(r'[^A-Za-z0-9]', '', str(value)).upper()


def find_col(df, candidates, label=""):
    """Return first matching column (case-insensitive). Prints options if not found."""
    lower_map = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    print(f"\n  ERROR: Could not find {label or candidates} column.")
    print(f"  Available columns: {list(df.columns)}")
    raise SystemExit(1)


# ── Version existing output ────────────────────────────────────────────────────
if os.path.exists(OUTPUT_FILE):
    mtime     = os.path.getmtime(OUTPUT_FILE)
    date_str  = datetime.fromtimestamp(mtime).strftime("%m.%d.%Y")
    archive   = os.path.join(DATA_DIR, f"BAT {date_str}.xlsx")
    shutil.copy2(OUTPUT_FILE, archive)
    print(f"  Archived previous report -> BAT {date_str}.xlsx")


# ── Load source files ──────────────────────────────────────────────────────────
print("\nLoading source files...")
detail_df = pd.read_csv(DETAIL_FILE)
policy_df = pd.read_csv(POLICY_FILE)
calls_df  = pd.read_csv(CALLS_FILE)

print(f"  Detail Report         : {len(detail_df):,} rows | columns: {list(detail_df.columns)}")
print(f"  PolicyIndividualsGroups: {len(policy_df):,} rows | columns: {list(policy_df.columns)}")
print(f"  Calls                 : {len(calls_df):,} rows | columns: {list(calls_df.columns)}")


# ── Identify columns ───────────────────────────────────────────────────────────
mbi_col      = find_col(detail_df, ['Mbi','MBI','mbi'],                      "MBI (Detail)")
advocate_col = find_col(detail_df, ['Advocate','advocate'],                   "Advocate")
agent_col    = find_col(detail_df, ['Writing Agent Name','Writing Agent',
                                     'WritingAgent','Agent Name','AgentName',
                                     'Agent'],                                 "Writing Agent")

policy_mbi_col = find_col(policy_df, ['Mbi','MBI','mbi'],                    "MBI (PolicyIndividualsGroups)")

# Auto-detect phone columns in PolicyIndividualsGroups
phone_kws = ['phone','number','cell','mobile','contact','tel']
phone_cols_policy = [c for c in policy_df.columns
                     if any(kw in c.lower() for kw in phone_kws)]
if not phone_cols_policy:
    print(f"\n  ERROR: No phone columns found in PolicyIndividualsGroups.")
    print(f"  Available columns: {list(policy_df.columns)}")
    raise SystemExit(1)

calls_phone_col = find_col(calls_df,
    ['phone_number_dialed','phone_number','Phone','PhoneNumber',
     'Phone Number','Number','Called','To','From','Caller','ANI','DNIS'],
    "Phone (Calls)")

print(f"\n  Mapped columns:")
print(f"    Detail            -> MBI='{mbi_col}'  Advocate='{advocate_col}'  Agent='{agent_col}'")
print(f"    PolicyIndivGroups -> MBI='{policy_mbi_col}'  Phones={phone_cols_policy}")
print(f"    Calls             -> Phone='{calls_phone_col}'")


# ── Build sets for attempted lookup ───────────────────────────────────────────
calls_phones = set(calls_df[calls_phone_col].dropna().apply(normalize))
calls_phones.discard("")

policy_df['_mbi_norm'] = policy_df[policy_mbi_col].apply(normalize)
mbi_phones: dict[str, set] = {}
for _, row in policy_df.iterrows():
    mbi = row['_mbi_norm']
    if not mbi:
        continue
    phones = {normalize(row[c]) for c in phone_cols_policy if not pd.isna(row[c])}
    phones.discard("")
    mbi_phones.setdefault(mbi, set()).update(phones)

def was_attempted(mbi_norm: str) -> bool:
    return bool(mbi_phones.get(mbi_norm, set()) & calls_phones)


# ── Aggregate Detail Report ────────────────────────────────────────────────────
detail_df['_mbi_norm']  = detail_df[mbi_col].apply(normalize)
detail_df['_completed'] = (
    detail_df[advocate_col].notna() &
    (detail_df[advocate_col].astype(str).str.strip() != '')
)

# One row per (agent, mbi) -- completed = True if ANY row for that MBI is completed
agent_mbi = (
    detail_df
    .groupby([agent_col, '_mbi_norm'], as_index=False)
    .agg(completed=('_completed', 'max'))
)
# Attempted = not yet completed AND phone appeared in Calls
agent_mbi['attempted'] = (
    ~agent_mbi['completed'] &
    agent_mbi['_mbi_norm'].apply(was_attempted)
)


# ── Per-agent metrics ──────────────────────────────────────────────────────────
def metrics(df):
    n         = len(df)
    completed = int(df['completed'].sum())
    attempted = int(df['attempted'].sum())
    return pd.Series({
        'Total Enrollments':  n,
        'Total Completed':    completed,
        'Total Completed %':  round(completed / n * 100, 1) if n else 0,
        'Total Attempted':    attempted,
        'Total Attempted %':  round(attempted / n * 100, 1) if n else 0,
        'Total Overall %':    round((completed + attempted) / n * 100, 1) if n else 0,
    })

agent_summary = (
    agent_mbi
    .groupby(agent_col)
    .apply(metrics)
    .reset_index()
    .rename(columns={agent_col: 'Writing Agent'})
    .sort_values('Writing Agent')
    .reset_index(drop=True)
)


# ── Overall row (unique MBIs across all agents) ────────────────────────────────
overall_mbi = (
    detail_df
    .groupby('_mbi_norm', as_index=False)
    .agg(completed=('_completed', 'max'))
)
overall_mbi['attempted'] = (
    ~overall_mbi['completed'] &
    overall_mbi['_mbi_norm'].apply(was_attempted)
)

n          = len(overall_mbi)
completed  = int(overall_mbi['completed'].sum())
attempted  = int(overall_mbi['attempted'].sum())

overall_row = pd.DataFrame([{
    'Writing Agent':    'OVERALL',
    'Total Enrollments': n,
    'Total Completed':   completed,
    'Total Completed %': round(completed / n * 100, 1) if n else 0,
    'Total Attempted':   attempted,
    'Total Attempted %': round(attempted / n * 100, 1) if n else 0,
    'Total Overall %':   round((completed + attempted) / n * 100, 1) if n else 0,
}])

final_df = pd.concat([overall_row, agent_summary], ignore_index=True)


# ── Write formatted Excel ──────────────────────────────────────────────────────
print("\nWriting Benefit Activation Read Out.xlsx ...")

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    final_df.to_excel(writer, sheet_name='Benefit Activation', index=False)
    ws = writer.sheets['Benefit Activation']

    # Styles
    hdr_fill     = PatternFill("solid", fgColor="1F3864")
    hdr_font     = Font(bold=True, color="FFFFFF", size=11)
    overall_fill = PatternFill("solid", fgColor="BDD7EE")
    overall_font = Font(bold=True, size=11)
    alt_fill     = PatternFill("solid", fgColor="DEEAF1")
    white_fill   = PatternFill("solid", fgColor="FFFFFF")

    # Header row
    for cell in ws[1]:
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 20

    # Overall row (row 2)
    for cell in ws[2]:
        cell.fill      = overall_fill
        cell.font      = overall_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row=2, column=1).alignment = Alignment(horizontal='left', vertical='center')

    # Agent rows
    for row_idx in range(3, ws.max_row + 1):
        fill = alt_fill if row_idx % 2 == 0 else white_fill
        for cell in ws[row_idx]:
            cell.fill      = fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal='left', vertical='center')

    # Percent formatting
    pct_cols = [i + 1 for i, col in enumerate(final_df.columns) if '%' in col]
    for col_idx in pct_cols:
        for row_idx in range(2, ws.max_row + 1):
            ws.cell(row=row_idx, column=col_idx).number_format = '0.0"%"'

    # Column widths
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 4

    ws.freeze_panes = "B2"

print(f"\n  Done. Report saved to:")
print(f"  {OUTPUT_FILE}")
print(f"\n  Summary:")
print(f"    Total unique enrollments : {n:,}")
print(f"    Total completed          : {completed:,}  ({round(completed/n*100,1) if n else 0}%)")
print(f"    Total attempted          : {attempted:,}  ({round(attempted/n*100,1) if n else 0}%)")
print(f"    Writing Agents           : {len(agent_summary)}")
