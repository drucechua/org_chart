import math
import re
import pandas as pd
from graphviz import Digraph
from collections import defaultdict

# -------------------------------------------
# CONFIG
# -------------------------------------------
INPUT_FILE = "ideal_final_output.xlsx"
SHEET_NAME = 0
OUTPUT_FILE = "org_chart_dept_clusters"   # org_chart_dept_clusters.png
RANKDIR = "TB"                            # vertical
FONT = "Helvetica"

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

if "Organization Name" not in df.columns:
    df["Organization Name"] = "Unknown"

df["Organization Name"] = (
    df["Organization Name"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
    .replace({"": "Unknown"})
)

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
    """Compact, readable node label: Name + Title."""
    name = row.get("Name", "")
    title = row.get("Line Detail 1", "")
    title = title.strip() if isinstance(title, str) else ""
    if title:
        return f"{name}\n{title}"
    return name

def safe_name(s: str) -> str:
    """Make a string safe for use as a Graphviz ID."""
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_") or "cluster"

# roots = no manager
roots = df[df["Reports To"].apply(is_null)]["Unique Identifier"].tolist()

# lookups
id_to_row = {row["Unique Identifier"]: row for _, row in df.iterrows()}
id_to_label = {uid: build_label(row) for uid, row in id_to_row.items()}

# group people by department
org_to_ids = defaultdict(list)
for uid, row in id_to_row.items():
    org = row.get("Organization Name", "Unknown")
    org_to_ids[org].append(uid)

# -------------------------------------------
# COLOR PALETTE (soft, not shouting)
# -------------------------------------------
# pastel-ish department colors
palette = [
    "#E3F2FD",  # light blue
    "#FFF3E0",  # light orange
    "#E8F5E9",  # light green
    "#F3E5F5",  # light purple
    "#E0F7FA",  # light cyan
    "#FBE9E7",  # light coral
    "#FFFDE7",  # light yellow
]

org_names = sorted(org_to_ids.keys())
org_to_color = {
    org: palette[i % len(palette)] for i, org in enumerate(org_names)
}

# -------------------------------------------
# CREATE GRAPH
# -------------------------------------------
dot = Digraph(comment="Org Chart (Dept Clusters)", format="png")

dot.graph_attr.update(
    rankdir=RANKDIR,
    splines="ortho",
    fontsize="11",
    labelloc="t",
    label="Org Chart",
    pad="0.2",
    margin="0.1",
    nodesep="0.3",
    ranksep="0.5",
    ratio="compress",
    bgcolor="white",
)

dot.node_attr.update(
    shape="box",
    style="rounded,filled",
    fillcolor="white",
    color="#555555",
    fontname=FONT,
    fontsize="9",
    margin="0.12,0.06",
)

dot.edge_attr.update(
    color="#888888",
    arrowsize="0.7",
)

# -------------------------------------------
# NODES: ADD DEPARTMENTS AS CLUSTERS
# -------------------------------------------
for org in org_names:
    dept_nodes = org_to_ids[org]
    cluster_name = f"cluster_{safe_name(org)}"
    dept_color = org_to_color[org]

    with dot.subgraph(name=cluster_name) as c:
        # Department frame
        c.attr(
            label=org,
            style="rounded,filled",
            color=dept_color,     # frame color
            fillcolor=dept_color, # soft background
            penwidth="1.4",
            fontsize="10",
            fontname=FONT,
        )

        # Nodes inside the department
        c.node_attr.update(
            style="rounded,filled",
            fillcolor="white",    # keep nodes neutral
            color="#555555",
            fontname=FONT,
            fontsize="9",
        )

        for uid in dept_nodes:
            label = id_to_label[uid]
            if uid in roots:
                # Top person(s) in org – slightly emphasized
                c.node(
                    uid,
                    label=label,
                    style="rounded,filled,bold",
                    penwidth="1.5",
                )
            else:
                c.node(uid, label=label)

# -------------------------------------------
# EDGES: TRUE REPORTING LINES
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
