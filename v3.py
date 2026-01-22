import math
import pandas as pd
from graphviz import Digraph

# -------------------------------------------
# CONFIG
# -------------------------------------------
INPUT_FILE = "ideal_final_output.xlsx"
SHEET_NAME = 0
OUTPUT_FILE = "org_chart_all"   # org_chart_all.png
RANKDIR = "TB"                  # vertical org chart

# -------------------------------------------
# LOAD DATA
# -------------------------------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

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
    """Compact label: Name + Title."""
    name = row.get("Name", "")
    title = row.get("Line Detail 1", "")
    title = title.strip() if isinstance(title, str) else ""
    if title:
        return f"{name}\n{title}"
    return name

# roots = no manager
roots = df[df["Reports To"].apply(is_null)]["Unique Identifier"].tolist()

id_to_row = {row["Unique Identifier"]: row for _, row in df.iterrows()}
id_to_label = {uid: build_label(row) for uid, row in id_to_row.items()}

# -------------------------------------------
# GRAPHVIZ (PNG, compact spacing)
# -------------------------------------------
dot = Digraph(comment="Org Chart (All Staff)", format="png")

dot.graph_attr.update(
    rankdir=RANKDIR,
    splines="ortho",
    fontsize="10",
    labelloc="t",
    label="Org Chart",
    pad="0.1",
    margin="0.05",
    nodesep="0.25",   # tighter horizontally
    ranksep="0.4",    # tighter vertically
    ratio="compress",
)

dot.node_attr.update(
    shape="box",
    style="rounded,filled",
    fillcolor="#f9f9f9",
    color="#555555",
    fontname="Helvetica",
    fontsize="9",
    margin="0.12,0.06",
)

dot.edge_attr.update(
    color="#888888",
    arrowsize="0.7",
)

# -------------------------------------------
# NODES (EVERYONE)
# -------------------------------------------
for uid, label in id_to_label.items():
    if uid in roots:
        dot.node(uid, label=label, fillcolor="#e3f2fd",
                 style="rounded,filled,bold", penwidth="1.3")
    else:
        dot.node(uid, label=label)

# -------------------------------------------
# EDGES (TRUE REPORTING LINES)
# -------------------------------------------
for _, row in df.iterrows():
    uid = row["Unique Identifier"]
    manager_id = row["Reports To"]

    if is_null(manager_id):
        continue

    manager_id = str(manager_id)

    if manager_id in id_to_label and uid in id_to_label:
        dot.edge(manager_id, uid)

# -------------------------------------------
# RENDER
# -------------------------------------------
output_path = dot.render(filename=OUTPUT_FILE, cleanup=True)
print(f"PNG org chart generated: {output_path}")
