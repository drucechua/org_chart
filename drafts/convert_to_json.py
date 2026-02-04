import pandas as pd
import math
import json

def is_null(x):
    return (
        x is None or
        (isinstance(x, float) and math.isnan(x)) or
        (isinstance(x, str) and x.strip() == "")
    )

def build_tree(df):
    df["Unique Identifier"] = df["Unique Identifier"].astype(str)
    df["Reports To"] = df["Reports To"].astype(str)

    rows = {r["Unique Identifier"]: r for _, r in df.iterrows()}
    children_map = {uid: [] for uid in rows}

    roots = []

    for uid, r in rows.items():
        mgr = r.get("Reports To")
        if is_null(mgr) or mgr not in children_map:
            roots.append(uid)
        else:
            children_map[mgr].append(uid)

    def to_node(uid):
        r = rows[uid]
        return {
            "id": uid,
            "name": r["Name"],
            "title": r.get("Line Detail 1", ""),
            "department": r.get("Organization Name", ""),
            "children": [to_node(c) for c in children_map[uid]]
        }

    return [to_node(root) for root in roots]

# Load Excel and convert to JSON
df = pd.read_excel("ideal_final_output.xlsx", sheet_name=0)
tree = build_tree(df)

with open("org_data.json", "w") as f:
    json.dump(tree, f, indent=2)

print("Saved org_data.json")
