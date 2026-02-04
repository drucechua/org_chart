import math
from collections import defaultdict, deque

import pandas as pd
from graphviz import Digraph

# -------------------------------------------
# CONFIG
# -------------------------------------------
INPUT_FILE = "ideal_final_output.xlsx"
SHEET_NAME = 0

# Single overall org-chart output
OUTPUT_FILE = "org_chart_hr_full"   # will produce org_chart_hr_full.png

RANKDIR = "TB"  # top-bottom

# Column used for grouping into functions / teams
GROUP_BY_COLUMN = "Organization Name"

# Rule: max peers per row (per level, per function)
# Applied ONLY to nodes that are NOT inside any team-leader subtree
MAX_PEERS_PER_ROW = 4

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

if GROUP_BY_COLUMN not in df.columns:
    raise ValueError(f"Column '{GROUP_BY_COLUMN}' not found in data.")


# -------------------------------------------
# HELPERS
# -------------------------------------------
def is_null(x):
    return (
        x is None
        or (isinstance(x, float) and math.isnan(x))
        or (isinstance(x, str) and (x.strip() == "" or x.strip().lower() == "nan"))
    )


def build_label(row):
    """Position-first label: Title on first line, Name on second."""
    title = row.get("Line Detail 1", "")
    name = row.get("Name", "")
    title = title.strip() if isinstance(title, str) else ""
    name = name.strip() if isinstance(name, str) else ""
    if title and name:
        return f"{title}\n{name}"
    elif title:
        return title
    return name


def norm_function(row):
    """
    Normalize the organization/function name:
    - Take everything before '('
    - Strip whitespace
    """
    val = row.get(GROUP_BY_COLUMN)
    if is_null(val):
        return "Other"
    s = str(val)
    if "(" in s:
        s = s.split("(", 1)[0]
    s = s.strip()
    return s or "Other"


# Roots = positions with no manager
roots = df[df["Reports To"].apply(is_null)]["Unique Identifier"].tolist()

# Lookups
id_to_row = {row["Unique Identifier"]: row for _, row in df.iterrows()}
id_to_label = {uid: build_label(row) for uid, row in id_to_row.items()}

# Parent / children maps
parent = {}
children = defaultdict(list)
for _, row in df.iterrows():
    uid = row["Unique Identifier"]
    manager_id = row["Reports To"]
    if not is_null(manager_id):
        manager_id = str(manager_id)
        parent[uid] = manager_id
        children[manager_id].append(uid)


# -------------------------------------------
# LEVELS (true hierarchy depth)
# -------------------------------------------
levels = {}

for root in roots:
    if root not in id_to_row:
        continue
    queue = deque([(root, 0)])
    while queue:
        node, lvl = queue.popleft()
        if node in levels and levels[node] <= lvl:
            continue
        levels[node] = lvl
        for ch in children.get(node, []):
            queue.append((ch, lvl + 1))


# -------------------------------------------
# TEAM LEADERS & STACKABLE NODES
# -------------------------------------------
def subtree_of(root_uid):
    """Return all nodes in the subtree rooted at root_uid (root + descendants)."""
    result = set()
    queue = deque([root_uid])
    while queue:
        u = queue.popleft()
        if u in result:
            continue
        result.add(u)
        for ch in children.get(u, []):
            queue.append(ch)
    return result


# Team leader rule:
# If a person has an Organization Name and is NOT the root (level > 0),
# they are a team leader. Everyone under them can be "stacked" (i.e., we
# don't apply peer-banding to those subtrees; they can form vertical branches).
team_leads = [
    uid
    for uid, row in id_to_row.items()
    if not is_null(row.get(GROUP_BY_COLUMN)) and levels.get(uid, 0) > 0
]

stackable_nodes = set()
for lead in team_leads:
    stackable_nodes |= subtree_of(lead)


# -------------------------------------------
# LAYOUT LEVELS (bands for non-stackable nodes only)
# -------------------------------------------
layout_levels = {}

# (true level, function) -> list of uids that are NOT in any team subtree
level_func_to_nodes = defaultdict(list)

for _, row in df.iterrows():
    uid = row["Unique Identifier"]
    if uid in stackable_nodes:
        continue  # do not band stackable nodes; they get natural tree layout
    lvl = levels.get(uid)
    if lvl is None:
        continue
    func = norm_function(row)
    level_func_to_nodes[(lvl, func)].append(uid)

# For each (level, function), break peers into bands of up to MAX_PEERS_PER_ROW
for (lvl, func), uids in level_func_to_nodes.items():
    band_index = 0
    for i in range(0, len(uids), MAX_PEERS_PER_ROW):
        band = uids[i : i + MAX_PEERS_PER_ROW]
        layout_level_value = lvl * 10 + band_index  # artificial layout level
        for uid in band:
            layout_levels[uid] = layout_level_value
        band_index += 1


# -------------------------------------------
# PALETTE FOR FUNCTIONS
# -------------------------------------------
PALETTE = [
    "#e3f2fd",  # blue-ish
    "#e8f5e9",  # green-ish
    "#fff3e0",  # orange-ish
    "#f3e5f5",  # purple-ish
    "#e0f7fa",  # teal-ish
    "#fce4ec",  # pink-ish
    "#f9fbe7",  # lime-ish
]

function_colors = {}


def get_function_color(func_name):
    if func_name not in function_colors:
        idx = len(function_colors) % len(PALETTE)
        function_colors[func_name] = PALETTE[idx]
    return function_colors[func_name]


# -------------------------------------------
# GRAPH CREATION
# -------------------------------------------
def add_node(dot_obj, uid, use_function_color=True):
    row = id_to_row[uid]
    label = id_to_label[uid]
    func = norm_function(row)
    lvl = levels.get(uid, 99)

    attrs = {}
    if use_function_color:
        attrs["fillcolor"] = get_function_color(func)

    # Emphasize top levels visually
    if lvl == 0:
        attrs.update(
            {
                "style": "rounded,filled,bold",
                "penwidth": "1.6",
                "fontsize": "11",
            }
        )
    elif lvl == 1:
        attrs.update(
            {
                "style": "rounded,filled",
                "penwidth": "1.3",
                "fontsize": "10",
            }
        )

    dot_obj.node(uid, label=label, **attrs)


def make_graph(
    included_nodes,
    filename,
    title="Org Chart – Office of Human Resources",
    use_clusters=True,
):
    """Create and render a Graphviz organizational chart."""
    dot = Digraph(comment=title, format="png")

    dot.graph_attr.update(
        rankdir=RANKDIR,
        splines="ortho",
        fontsize="11",
        labelloc="t",
        label=title,
        pad="0.2",
        margin="0.1",
        nodesep="0.35",
        ranksep="0.6",
        newrank="true",
        center="true",
    )

    dot.node_attr.update(
        shape="box",
        style="rounded,filled",
        fillcolor="#f9f9f9",
        color="#555555",
        fontname="Helvetica",
        fontsize="9",
        margin="0.14,0.08",
    )

    dot.edge_attr.update(
        color="#888888",
        arrowsize="0.7",
    )

    included_nodes = set(included_nodes)

    # Group by function for containers/columns
    if use_clusters:
        func_to_nodes = defaultdict(list)
        for uid in included_nodes:
            func = norm_function(id_to_row[uid])
            func_to_nodes[func].append(uid)

        for i, (func, members) in enumerate(sorted(func_to_nodes.items())):
            with dot.subgraph(name=f"cluster_{i}") as c:
                c.attr(
                    label=func,  # function name only
                    style="rounded,dashed",
                    color="#cccccc",
                    fontsize="10",
                    fontname="Helvetica",
                )
                for uid in members:
                    add_node(c, uid)
    else:
        for uid in included_nodes:
            add_node(dot, uid)

    # Rank bands for NON-stackable nodes only (respecting MAX_PEERS_PER_ROW)
    level_to_nodes = defaultdict(list)
    for uid in included_nodes:
        if uid not in layout_levels:
            continue  # stackable nodes: no artificial banding here
        ll = layout_levels[uid]
        level_to_nodes[ll].append(uid)

    for ll, nodes_at_ll in level_to_nodes.items():
        with dot.subgraph() as s:
            s.attr(rank="same")
            for uid in nodes_at_ll:
                s.node(uid)

    # Edges (true reporting lines)
    for uid in included_nodes:
        m = parent.get(uid)
        if m is None:
            continue
        if m not in included_nodes:
            continue
        dot.edge(m, uid)

    out_path = dot.render(filename=filename, cleanup=True)
    print(f"Generated: {out_path}")


# -------------------------------------------
# SINGLE OVERALL ORG CHART
# -------------------------------------------
all_nodes = list(id_to_row.keys())

make_graph(
    included_nodes=all_nodes,
    filename=OUTPUT_FILE,
    title="Org Chart – Office of Human Resources",
    use_clusters=True,
)
