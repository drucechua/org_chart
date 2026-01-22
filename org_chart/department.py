import pandas as pd
from graphviz import Digraph
import math
import re

# -------------------------------------------
# CONFIG
# -------------------------------------------
INPUT_FILE = "ideal_final_output.xlsx"
SHEET_NAME = 0        # first sheet; change if needed
OUTPUT_PREFIX = "org_chart"  # will create org_chart_<dept>.png
RANKDIR = "TB"        # "TB" = top-bottom, "LR" = left-right

# -------------------------------------------
# LOAD DATA
# -------------------------------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

# Expected columns:
# - Unique Identifier
# - Name
# - Reports To
# - Line Detail 1 (title)
# - Line Detail 2 (location)
# - Line Detail 3
# - Organization Name

# Clean up types
df["Unique Identifier"] = df["Unique Identifier"].astype(str)
df["Name"] = df["Name"].astype(str).str.strip()
if "Reports To" in df.columns:
    df["Reports To"] = df["Reports To"].astype(str).replace({"nan": None})
else:
    df["Reports To"] = None

# -------------------------------------------
# HELPERS
# -------------------------------------------
def is_null(x):
    return (
        x is None
        or (isinstance(x, float) and math.isnan(x))
        or (isinstance(x, str) and x.strip() == "")
    )

def build_label(row):
    """
    Line 1: Name
    Line 2: Title (Line Detail 1)
    """
    name = row.get("Name", "")
    title = row.get("Line Detail 1", "")

    lines = [name]
    if isinstance(title, str) and title.strip():
        lines.append(title.strip())

    return "\n".join(lines)

def extract_dept_name(org_name):
    """
    From 'HR Planning  (Moussoux, Florence)' -> 'HR Planning'
    From 'Relocation Services' -> 'Relocation Services'
    """
    if pd.isna(org_name):
        return None
    s = str(org_name).strip()
    # split at first '(' if present
    if "(" in s:
        s = s.split("(", 1)[0]
    return s.strip()

def safe_filename(s):
    """
    Turn department name into a safe filename suffix.
    """
    if not s:
        s = "Unknown"
    s = re.sub(r"[^A-Za-z0-9]+", "_", s)
    s = s.strip("_")
    return s or "Unknown"

# -------------------------------------------
# GLOBAL LOOKUP: LABELS & REPORTING TREE
# -------------------------------------------
# Node labels for everyone
id_to_label = {
    row["Unique Identifier"]: build_label(row)
    for _, row in df.iterrows()
}

# Build adjacency: manager -> list of direct reports
manager_to_reports = {}
for _, row in df.iterrows():
    uid = str(row["Unique Identifier"])
    mgr = row["Reports To"]
    if is_null(mgr):
        continue
    mgr = str(mgr)
    manager_to_reports.setdefault(mgr, []).append(uid)

def get_subtree_nodes(root_id):
    seen = set()
    stack = [root_id]

    while stack:
        current = stack.pop()
        if current in seen:
            continue

        seen.add(current)

        for child in manager_to_reports.get(current, []):
            if child not in seen:
                stack.append(child)

    return seen
