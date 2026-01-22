import pandas as pd
import math
import json
from collections import defaultdict

# -------------------------------------------
# CONFIG
# -------------------------------------------
INPUT_FILE = "ideal_final_output.xlsx"
SHEET_NAME = 0              # first sheet
OUTPUT_HTML = "org_chart.html"

# Columns (adjust if your file uses different names)
COL_ID = "Unique Identifier"
COL_NAME = "Name"
COL_REPORTS_TO = "Reports To"
COL_TITLE = "Line Detail 1"
COL_ORG = "Organization Name"

# -------------------------------------------
# LOAD DATA
# -------------------------------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

# Clean up types
df[COL_ID] = df[COL_ID].astype(str)
df[COL_NAME] = df[COL_NAME].astype(str).str.strip()

if COL_REPORTS_TO in df.columns:
    df[COL_REPORTS_TO] = df[COL_REPORTS_TO].astype(str).replace({"nan": None})
else:
    df[COL_REPORTS_TO] = None

# Ensure optional columns exist
for col in [COL_TITLE, COL_ORG]:
    if col not in df.columns:
        df[col] = ""

# -------------------------------------------
# HELPER: null check
# -------------------------------------------
def is_null(x):
    if x is None:
        return True
    if isinstance(x, float) and math.isnan(x):
        return True
    if isinstance(x, str) and x.strip() == "" or x.strip().lower() == "nan":
        return True
    return False

# -------------------------------------------
# BUILD BASIC LOOKUP
# -------------------------------------------
nodes = {}

for _, row in df.iterrows():
    uid = row[COL_ID]
    nodes[uid] = {
        "id": uid,
        "name": row.get(COL_NAME, "") or "",
        "title": row.get(COL_TITLE, "") or "",
        "org": row.get(COL_ORG, "") or "",
        "children": []
    }

# -------------------------------------------
# BUILD PARENT → CHILD RELATIONSHIPS
# -------------------------------------------
# Track roots (no manager)
parent_map = {}
children_map = defaultdict(list)

for _, row in df.iterrows():
    uid = row[COL_ID]
    manager_id = row[COL_REPORTS_TO]

    if is_null(manager_id):
        parent_map[uid] = None
    else:
        manager_id = str(manager_id)
        parent_map[uid] = manager_id
        children_map[manager_id].append(uid)

# Attach children into node dicts
for parent_id, child_ids in children_map.items():
    if parent_id in nodes:
        nodes[parent_id]["children"] = [nodes[cid] for cid in child_ids if cid in nodes]

# -------------------------------------------
# FIND ROOTS
# -------------------------------------------
roots = [uid for uid, pid in parent_map.items() if pid is None]

print(f"[INFO] Detected {len(roots)} root node(s): {roots}")
if len(roots) == 0:
    raise RuntimeError("No root nodes detected – cannot build org chart.")
elif len(roots) > 1:
    print("[WARN] Multiple roots detected. The chart will have multiple top-level trees.")

# For now we assume single main root; if multiple, we wrap them under a virtual root
if len(roots) == 1:
    root_id = roots[0]
    root_node = nodes[root_id]
else:
    # virtual root to hold multiple trees – OrgChart can show this as a dummy node
    root_node = {
        "id": "VIRTUAL_ROOT",
        "name": "Organization",
        "title": "",
        "org": "",
        "children": [nodes[rid] for rid in roots]
    }

# -------------------------------------------
# SERIALIZE HIERARCHY TO JSON
# -------------------------------------------
hierarchy_json = json.dumps(root_node, indent=2)

# -------------------------------------------
# BUILD HTML WITH ORGCHART INTEGRATION
# -------------------------------------------
# Using OrgChart by dabeng via CDN. If these URLs change in the future,
# update them in the <head> section.
html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>Org Chart</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />

  <!-- OrgChart CSS -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/orgchart@3.8.0/dist/css/jquery.orgchart.min.css">

  <!-- Basic page styling -->
  <style>
    html, body {{
      margin: 0;
      padding: 0;
      height: 100%;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
      background: #f5f5f7;
    }}
    #chart-container {{
      width: 100%;
      height: 100vh;
      overflow: auto;
      background: linear-gradient(180deg, #f5f5f7 0%, #ffffff 40%);
    }}
    .orgchart {{
      background: transparent !important;
    }}
    .orgchart .node {{
      border-radius: 10px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.08);
      border: 1px solid #d0d7de;
      background: #ffffff;
      padding: 6px 10px;
    }}
    .org-node-name {{
      font-weight: 600;
      font-size: 13px;
      margin-bottom: 2px;
      color: #111827;
    }}
    .org-node-title {{
      font-size: 11px;
      color: #4b5563;
    }}
    .org-node-org {{
      font-size: 10px;
      color: #6b7280;
      margin-top: 3px;
    }}
    .orgchart .node.focused {{
      border-color: #2563eb;
      box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.2);
    }}
  </style>

  <!-- jQuery -->
  <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>

  <!-- OrgChart JS -->
  <script src="https://cdn.jsdelivr.net/npm/orgchart@3.8.0/dist/js/jquery.orgchart.min.js"></script>
</head>
<body>
  <div id="chart-container"></div>

  <script>
    // Hierarchical data from your Python pipeline
    var orgData = {hierarchy_json};

    // Initialize OrgChart once DOM is ready
    $(function() {{
      var oc = $('#chart-container').orgchart({{
        data: orgData,
        nodeTitle: 'name',
        nodeContent: 'title',   // fallback for built-in; we mainly use nodeTemplate below
        pan: true,
        zoom: true,
        draggable: true,
        direction: 't2b',       // top-to-bottom
        verticalLevel: 3,       // mix vertical/horizontal to help large charts
        visibleLevel: 3,        // don't show absolutely everything at first
        exportButton: false,
        nodeTemplate: function(data) {{
          var orgLine = data.org ? '<div class="org-node-org">' + data.org + '</div>' : '';
          var titleLine = data.title ? '<div class="org-node-title">' + data.title + '</div>' : '';
          return (
            '<div class="org-node-name">' + (data.name || '') + '</div>' +
            titleLine +
            orgLine
          );
        }},
        createNode: function($node, data) {{
          // Optional: highlight team leaders (have org)
          if (data.org) {{
            $node.addClass('team-leader');
          }}
          // Focus style on click
          $node.on('click', function() {{
            $('.orgchart .node').removeClass('focused');
            $(this).addClass('focused');
          }});
        }}
      }});
    }});
  </script>
</body>
</html>
"""

with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(html_template)

print(f"[INFO] OrgChart HTML generated: {OUTPUT_HTML}")
print("Open this file in a browser to view the interactive org chart.")
