import pandas as pd
from graphviz import Digraph

# ----------------------------
# LOAD DATA
# ----------------------------
df = pd.read_excel("ideal_final_output.xlsx")

df["Unique Identifier"] = df["Unique Identifier"].astype(str).str.strip()
df["Reports To"] = (
    df["Reports To"]
    .astype(str)
    .str.strip()
    .replace({"nan": None, "None": None, "": None})
)
df["Name"] = df["Name"].astype(str).str.strip()

# ----------------------------
# FIND FLORENCE
# ----------------------------
# Adjust the "Florence" string if needed to match your data
florence_row = df[df["Name"].str.contains("Florence", case=False)].iloc[0]

florence_id = florence_row["Unique Identifier"]
florence_name = florence_row["Name"]
florence_title = florence_row["Line Detail 1"]
florence_org = florence_row["Organization Name"]

# ----------------------------
# BUILD REPORTING MAP
# ----------------------------
manager_to_reports = {}

for _, row in df.iterrows():
    mgr = row["Reports To"]
    uid = row["Unique Identifier"]
    if mgr:
        manager_to_reports.setdefault(mgr, []).append(uid)

# ----------------------------
# GET SUBTREE (FLORENCE + ALL REPORTS)
# ----------------------------
subtree = set()
stack = [florence_id]

while stack:
    current = stack.pop()
    if current in subtree:
        continue
    subtree.add(current)
    for r in manager_to_reports.get(current, []):
        stack.append(r)

# ----------------------------
# BUILD GRAPH
# ----------------------------
dot = Digraph(format="png")
dot.attr(rankdir="TB")

# Header/description at top of the chart
dot.attr(
    label=f"Team of {florence_name}",
    labelloc="t",
    fontsize="12",
    fontname="Helvetica"
)

dot.attr(
    "node",
    shape="box",
    style="rounded,filled",
    fontname="Helvetica",
    fontsize="10"
)

# Nodes
for uid in subtree:
    row = df.loc[df["Unique Identifier"] == uid].iloc[0]
    name = row["Name"]
    title = row["Line Detail 1"]

    if uid == florence_id:
        label = f"{name}\n{title}\n{florence_org}"
        dot.node(uid, label=label, fillcolor="#e3f2fd")  # Florence highlighted
    else:
        label = f"{name}\n{title}"
        dot.node(uid, label=label, fillcolor="#f9f9f9")

# Edges (use actual manager relationships within this subtree)
for _, row in df.iterrows():
    uid = row["Unique Identifier"]
    mgr = row["Reports To"]
    if mgr and uid in subtree and mgr in subtree:
        dot.edge(mgr, uid)

# ----------------------------
# RENDER
# ----------------------------
dot.render("org_chart_Florence", cleanup=True)
print("Generated org_chart_Florence.png")
