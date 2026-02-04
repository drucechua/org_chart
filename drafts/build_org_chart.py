import pandas as pd
from graphviz import Digraph
import math

# -------------------------------------------
# CONFIG
# -------------------------------------------
INPUT_FILE = "ideal_final_output.xlsx"
SHEET_NAME = 0        # first sheet; change if needed
OUTPUT_FILE = "org_chart"  # will create org_chart.png (or .pdf)
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
# BUILD A LOOKUP FOR NODE LABELS
# -------------------------------------------
def build_label(row):
    """
    Build a nice multi-line label for each node:
    Line 1: Name
    Line 2: Title
    Line 3: Org (if available)
    """
    name = row.get("Name", "")
    title = row.get("Line Detail 1", "")
    # org = row.get("Organization Name", "")

    lines = [name]
    if isinstance(title, str) and title.strip():
        lines.append(title.strip())
    # if isinstance(org, str) and org.strip():
    #     lines.append(org.strip())

    return "\n".join(lines)

id_to_label = {
    row["Unique Identifier"]: build_label(row)
    for _, row in df.iterrows()
}

# -------------------------------------------
# IDENTIFY ROOT NODES (NO MANAGER)
# -------------------------------------------
# "Reports To" is None, NaN, or empty → root
def is_null(x):
    return x is None or (isinstance(x, float) and math.isnan(x)) or (isinstance(x, str) and x.strip() == "")

roots = df[df["Reports To"].apply(is_null)]["Unique Identifier"].tolist()

# -------------------------------------------
# CREATE GRAPHVIZ DIGRAPH
# -------------------------------------------
dot = Digraph(comment="HR Org Chart", format="png")
dot.attr(rankdir=RANKDIR)  # TB or LR
dot.attr(
    "node",
    shape="box",
    style="rounded,filled",
    fillcolor="#f9f9f9",
    color="#555555",
    fontname="Helvetica",
    fontsize="10"
)
dot.attr("edge", color="#888888", arrowsize="0.7")

# -------------------------------------------
# ADD NODES
# -------------------------------------------
for uid, label in id_to_label.items():
    # You could color root(s) differently if you want
    if uid in roots:
        dot.node(uid, label=label, fillcolor="#e3f2fd")  # light blue for top-level
    else:
        dot.node(uid, label=label)

# -------------------------------------------
# ADD EDGES (MANAGER → EMPLOYEE)
# -------------------------------------------
for _, row in df.iterrows():
    uid = row["Unique Identifier"]
    manager_id = row["Reports To"]

    if is_null(manager_id):
        continue  # root node
    manager_id = str(manager_id)

    # Only add edge if both nodes exist
    if manager_id in id_to_label and uid in id_to_label:
        dot.edge(manager_id, uid)

# -------------------------------------------
# RENDER TO FILE
# -------------------------------------------
output_path = dot.render(filename=OUTPUT_FILE, cleanup=True)
print(f"Org chart generated: {output_path}")
