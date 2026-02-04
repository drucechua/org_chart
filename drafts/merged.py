import pandas as pd
import math
import json
from collections import defaultdict
import os
import webbrowser
from typing import Optional

# -------------------------------------------
# CONFIG
# -------------------------------------------

# Raw input file from HR
RAW_INPUT_FILE = "Office of Human Resources  (AlNuaimi, Rashed).xlsx"
RAW_SHEET_NAME = "Org Chart"

# Optional: where to save the cleaned dataset
CLEANED_OUTPUT_FILE = "ideal_final_output.xlsx"

# Org chart output
OUTPUT_HTML = "org_chart_latest.html"

# Column names (matching your Excel)
COL_ID = "Unique Identifier"
COL_NAME = "Name"
COL_REPORTS_TO = "Reports To"
COL_TITLE = "Line Detail 1"
COL_ORG = "Organization Name"


# ===========================================
# PART 1: CLEANING / NORMALIZATION
# (your first script, wrapped as a function)
# ===========================================

def load_and_clean_org_data(
    input_path: str = RAW_INPUT_FILE,
    sheet_name: str = RAW_SHEET_NAME,
    output_path: Optional[str] = CLEANED_OUTPUT_FILE,
) -> pd.DataFrame:
    """
    Load raw HR org data, clean it into a person-level table,
    and optionally save it to Excel.
    Mirrors your original cleaning logic.
    """
    # -------------------------------------------
    # STEP 1 — LOAD ORIGINAL FILE
    # -------------------------------------------
    df = pd.read_excel(input_path, sheet_name=sheet_name)

    # -------------------------------------------
    # STEP 2 — REMOVE UNFILLED POSITIONS
    # And remove rows where Unique Identifier contains more than 5 digits
    # -------------------------------------------
    uid_series = df[COL_ID].astype(str)

    df = df[
        ~uid_series.str.contains("unfilled", case=False, na=False)  # remove unfilled rows
        & ~uid_series.str.contains(r"\d{6,}", regex=True)           # remove IDs with 6+ digits
    ].copy()

    # -------------------------------------------
    # STEP 3 — STANDARDIZE NAME FIELD
    # -------------------------------------------
    df[COL_NAME] = df[COL_NAME].astype(str).str.strip()

    # -------------------------------------------
    # STEP 4 — BUILD CANONICAL PERSON-LEVEL ID
    # Choose the first Unique Identifier for each Name
    # -------------------------------------------
    name_to_canonical_id = df.groupby(COL_NAME)[COL_ID].first().to_dict()

    # -------------------------------------------
    # STEP 5 — DROP DUPLICATES (ONE ROW PER PERSON)
    # -------------------------------------------
    df_unique = df.drop_duplicates(subset=[COL_NAME], keep="first").copy()
    df_unique[COL_ID] = df_unique[COL_NAME].map(name_to_canonical_id)

    # Ensure Reports To column exists
    if COL_REPORTS_TO not in df_unique.columns:
        df_unique[COL_REPORTS_TO] = pd.NA

    # -------------------------------------------
    # STEP 6 — NORMALIZE REPORTING LINES (NAME-BASED)
    # Convert old position-based IDs to canonical person IDs
    # -------------------------------------------
    def extract_name_from_id(raw_uid):
        if pd.isna(raw_uid):
            return None
        raw_uid = str(raw_uid)
        parts = raw_uid.split("_", 1)
        if len(parts) == 2:
            name = parts[1]
            return name.replace("_", " ")
        return None

    def normalize_reports_to(raw_uid):
        if pd.isna(raw_uid):
            return None
        name = extract_name_from_id(raw_uid)
        if name in name_to_canonical_id:
            return name_to_canonical_id[name]
        return None

    df_unique[COL_REPORTS_TO] = df_unique[COL_REPORTS_TO].apply(normalize_reports_to)

    # -------------------------------------------
    # STEP 7 — REMOVE SELF-REFERENCING REPORTS
    # -------------------------------------------
    mask_self = df_unique[COL_REPORTS_TO] == df_unique[COL_ID]
    df_unique.loc[mask_self, COL_REPORTS_TO] = pd.NA

    # -------------------------------------------
    # STEP 8 — OPTIONAL SAVE CLEANED DATA
    # -------------------------------------------
    if output_path is not None:
        df_unique.to_excel(output_path, index=False)
        print(f"[INFO] Cleaned data saved as: {output_path}")

    return df_unique


# ===========================================
# PART 2: ORG CHART GENERATION (OrgChart.js)
# (your second script, adjusted to trust cleaned data)
# ===========================================

def is_null(x):
    """
    Robust null check: handles None, pd.NA, NaN, and 'nan'/'none' strings.
    """
    if pd.isna(x):
        return True
    if isinstance(x, str):
        s = x.strip().lower()
        return s in ("", "nan", "none")
    return False


def is_leader_value(v):
    """Leader if Organization Name is non-empty."""
    if pd.isna(v):
        return False
    if isinstance(v, str):
        s = v.strip()
        return s != "" and s.lower() != "nan"
    return False


def build_org_chart(df: pd.DataFrame, output_html: str = OUTPUT_HTML):
    """
    Take a cleaned person-level DataFrame and generate an interactive
    org chart HTML (OrgChart.js).
    """
    # Work on a copy so we don't mutate the original df
    df = df.copy()

    # Basic normalization
    df[COL_ID] = df[COL_ID].astype(str)
    df[COL_NAME] = df[COL_NAME].astype(str).str.strip()

    # DO NOT cast Reports To to str globally — we want None / NA
    if COL_REPORTS_TO not in df.columns:
        df[COL_REPORTS_TO] = pd.NA

    # Ensure optional columns exist
    for col in [COL_TITLE, COL_ORG]:
        if col not in df.columns:
            df[col] = ""

    # -------------------------------------------
    # BUILD BASIC LOOKUP
    # -------------------------------------------
    nodes = {}

    for _, row in df.iterrows():
        uid = row[COL_ID]
        org_val = row.get(COL_ORG, "")

        # Normalize title safely to string
        raw_title = row.get(COL_TITLE, "")
        if isinstance(raw_title, str):
            full_title = raw_title.strip()
        else:
            full_title = ""

        leader_flag = is_leader_value(org_val)

        # Short title for display (first clause, then truncated)
        short_title = full_title or ""
        if isinstance(short_title, str) and "," in short_title:
            short_title = short_title.split(",")[0]
        short_title = short_title.strip()
        if len(short_title) > 40:
            short_title = short_title[:37].rstrip() + "…"

        # Role-based class for color-coding
        title_lower = full_title.lower()
        class_tokens = []

        # Trainees override everything else
        if "trainee" in title_lower:
            class_tokens.append("role-trainee")
        else:
            if leader_flag:
                class_tokens.append("role-leader")
            else:
                class_tokens.append("role-staff")

        class_name = " ".join(class_tokens)

        nodes[uid] = {
            "id": uid,
            "name": row.get(COL_NAME, "") or "",
            "title": full_title,         # full title (for tooltip)
            "shortTitle": short_title,   # concise title for node display
            "org": org_val or "",
            "children": [],
            "isLeader": leader_flag,     # used for styling + collapse logic
            "className": class_name,     # used for color-coding in OrgChart
        }

    # -------------------------------------------
    # BUILD PARENT → CHILD RELATIONSHIPS
    # -------------------------------------------
    parent_map = {}
    children_map = defaultdict(list)

    for _, row in df.iterrows():
        uid = row[COL_ID]
        manager_id = row[COL_REPORTS_TO]

        if is_null(manager_id):
            parent_map[uid] = None
        else:
            manager_id_str = str(manager_id)
            parent_map[uid] = manager_id_str
            children_map[manager_id_str].append(uid)

    # Attach children into node dicts
    for parent_id, child_ids in children_map.items():
        if parent_id in nodes:
            nodes[parent_id]["children"] = [
                nodes[cid] for cid in child_ids if cid in nodes
            ]

    # -------------------------------------------
    # FIND ROOTS
    # -------------------------------------------
    roots = [uid for uid, pid in parent_map.items() if pid is None]

    print(f"[INFO] Detected {len(roots)} root node(s): {roots}")
    if len(roots) == 0:
        raise RuntimeError("No root nodes detected – cannot build org chart.")
    elif len(roots) > 1:
        print("[WARN] Multiple roots detected. The chart will have multiple top-level trees.")

    # If multiple roots: wrap them under a virtual global root
    if len(roots) == 1:
        root_id = roots[0]
        root_node = nodes[root_id]
        # Mark CHRO as a distinct role for color-coding
        existing_cls = nodes[root_id].get("className", "")
        nodes[root_id]["className"] = (existing_cls + " role-chro").strip()
    else:
        root_id = "VIRTUAL_ROOT"
        root_node = {
            "id": root_id,
            "name": "Organization",
            "title": "",
            "shortTitle": "",
            "org": "",
            "children": [nodes[rid] for rid in roots],
            "isLeader": True,
            "isGroup": True,
        }

    # -------------------------------------------
    # INSERT GROUP NODES UNDER THE CHRO
    # -------------------------------------------
    group_ids = []
    if len(roots) == 1:
        original_children = root_node.get("children", [])

        # Define virtual group nodes
        group_leaders = {
            "id": "GROUP_LEADERS",
            "name": "LEADERSHIP & HEADS",
            "title": "Directors, Heads, Managers, Chiefs",
            "shortTitle": "LEADERSHIP & HEADS",
            "org": "",
            "children": [],
            "isLeader": True,
            "isGroup": True,
            "compact": False,
        }
        group_staff = {
            "id": "GROUP_STAFF",
            "name": "PROFESSIONAL STAFF",
            "title": "Coordinators, Specialists, Officers",
            "shortTitle": "PROFESSIONAL STAFF",
            "org": "",
            "children": [],
            "isLeader": True,
            "isGroup": True,
            "compact": False,
        }
        group_trainees = {
            "id": "GROUP_TRAINEES",
            "name": "TRAINEES & EARLY CAREER",
            "title": "Academic Operations Trainees & similar roles",
            "shortTitle": "TRAINEES & EARLY CAREER",
            "org": "",
            "children": [],
            "isLeader": True,
            "isGroup": True,
            "compact": True,  # compact layout still allowed for this branch
        }

        for child in original_children:
            title = (child.get("title") or "").lower()

            if "trainee" in title:
                group_trainees["children"].append(child)
            elif any(keyword in title for keyword in ["director", "head", "manager", "chief"]):
                group_leaders["children"].append(child)
            else:
                group_staff["children"].append(child)

        new_children = []
        for group in (group_leaders, group_staff, group_trainees):
            if group["children"]:
                new_children.append(group)
                group_ids.append(group["id"])

        if new_children:
            root_node["children"] = new_children
            print("[INFO] Applied CHRO-level grouping into virtual sections:",
                  [g["id"] for g in new_children])

    # -------------------------------------------
    # APPLY COLLAPSE LOGIC
    # -------------------------------------------
    default_expanded_group_id = "GROUP_LEADERS" if "GROUP_LEADERS" in group_ids else None

    def apply_collapse_flags(node, is_root=False, expanded_group_id=None):
        children = node.get("children", [])

        if children:
            if is_root:
                node["collapsed"] = False
            elif node.get("isGroup"):
                if expanded_group_id is not None and node.get("id") == expanded_group_id:
                    node["collapsed"] = False
                else:
                    node["collapsed"] = True
            else:
                node["collapsed"] = not node.get("isLeader", False)

            for child in children:
                apply_collapse_flags(child, is_root=False, expanded_group_id=expanded_group_id)

    apply_collapse_flags(root_node, is_root=True, expanded_group_id=default_expanded_group_id)

    # -------------------------------------------
    # SERIALIZE HIERARCHY TO JSON
    # -------------------------------------------
    hierarchy_json = json.dumps(root_node, indent=2)

    # -------------------------------------------
    # BUILD HTML WITH ORGCHART INTEGRATION
    # -------------------------------------------
    html_template = '''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Org Chart</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />

  <!-- OrgChart CSS -->
  <link rel="stylesheet"
        href="https://cdn.jsdelivr.net/npm/orgchart@3.8.0/dist/css/jquery.orgchart.min.css">

  <style>
    /* Page basics */
    html, body {
      margin: 0;
      padding: 0;
      height: 100%;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
      background: #f5f5f7;
    }
    #chart-container {
      width: 100%;
      height: 100vh;
      overflow: auto;
      background: linear-gradient(180deg, #f5f5f7 0%, #ffffff 40%);
      position: relative;
    }

    /* OrgChart basics */
    .orgchart {
      background: transparent !important;
    }
    .orgchart .nodes {
      align-items: flex-start; /* prevent vertical stretching of nodes */
    }

    /* Node visual style (uniform person cards) */
    .orgchart .node {
      border-radius: 10px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.08);
      border: 1px solid #d0d7de;
      background: #ffffff;
      padding: 8px;
      transition: box-shadow 0.18s ease, border-color 0.18s ease;
      box-sizing: border-box;
    }

    /* Standardize size for person nodes (exclude group nodes) */
    .orgchart .node:not(.group-node) {
      width: 190px;           /* fixed visual width */
      min-height: 72px;       /* fixed visual height */
      max-height: 72px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      text-align: center;
      box-sizing: border-box;
      padding: 8px;
    }

    /* Prevent text from increasing node height and truncate */
    .org-node-name,
    .org-node-title {
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .org-node-name {
      font-weight: 600;
      font-size: 13px;
      margin-bottom: 4px;
      color: #111827;
    }
    .org-node-title {
      font-size: 11px;
      color: #4b5563;
    }

    /* Leaders: light accent background (not faded) */
    .orgchart .node.team-leader {
      border-color: #2563eb;
      background: #eff6ff;
    }

    /* Focus style for clicked node */
    .orgchart .node.focused {
      border-color: #2563eb;
      box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.18);
    }

    /* Group nodes (section headers) */
    .orgchart .node.group-node {
      background: #111827;
      border-color: #111827;
      color: #f9fafb;
      box-shadow: 0 4px 10px rgba(15, 23, 42, 0.35);
      padding: 10px 14px;
    }
    .group-node .org-node-name {
      font-size: 14px;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      margin-bottom: 6px;
      color: #f9fafb;
    }
    .group-node .org-node-title {
      font-size: 11px;
      color: #e5e7eb;
    }

    /* ---------------------------------
       Color coding by role (via className)
       ----------------------------------*/

    /* CHRO: dark navy card */
    .orgchart .node.role-chro {
      background: #020617;
      border-color: #020617;
      color: #f9fafb;
    }
    .orgchart .node.role-chro .org-node-name {
      color: #f9fafb;
    }
    .orgchart .node.role-chro .org-node-title {
      color: #e5e7eb;
    }

    /* Leaders: blue accent (non-group) */
    .orgchart .node.role-leader:not(.group-node) {
      background: #eff6ff;
      border-color: #2563eb;
    }

    /* Regular staff: keep subtle grey card */
    .orgchart .node.role-staff:not(.group-node) {
      border-color: #d0d7de;
      background: #ffffff;
    }

    /* Trainees & early career: soft purple */
    .orgchart .node.role-trainee:not(.group-node) {
      background: #f5f3ff;
      border-color: #a855f7;
    }
    .orgchart .node.role-trainee .org-node-name {
      color: #312e81;
    }
    .orgchart .node.role-trainee .org-node-title {
      color: #4c1d95;
    }

    /* Level rhythm: give each level breathing room */
    .orgchart .level {
      padding-top: 14px;
      padding-bottom: 14px;
    }

    /* Soften connector lines a bit */
    .orgchart .lines .topLine,
    .orgchart .lines .leftLine,
    .orgchart .lines .rightLine,
    .orgchart .lines .downLine {
      border-color: #e6e9ee;
    }

    /* small hint bubble */
    .hint-bar {
      position: absolute;
      top: 10px;
      left: 16px;
      background: rgba(17, 24, 39, 0.78);
      color: #e5e7eb;
      padding: 6px 12px;
      border-radius: 999px;
      font-size: 12px;
      z-index: 10;
    }
  </style>
</head>
<body>
  <div class="hint-bar">Tip: Click a group or leader to expand their team.</div>
  <div id="chart-container"></div>

  <!-- jQuery + OrgChart -->
  <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/orgchart@3.8.0/dist/js/jquery.orgchart.min.js"></script>

  <script>
    // Data injected from Python
    var orgData = __ORG_DATA__;

    $(function() {
      var $container = $('#chart-container');

      function getChart() {
        return $container.find('.orgchart');
      }

      // Center whole chart (so global root is roughly centered)
      function centerChart() {
        var $chart = getChart();
        if (!$chart.length) return;

        var chartWidth = $chart.outerWidth();
        var containerWidth = $container.width();

        if (chartWidth > containerWidth) {
          $container.scrollLeft((chartWidth - containerWidth) / 2);
        } else {
          $container.scrollLeft(0);
        }
      }

      // Center horizontally on a specific node
      function centerOnNode($node) {
        var $chart = getChart();
        if (!$chart.length || !$node.length) return;

        var chartOffset = $chart.offset();
        var nodeOffset = $node.offset();
        var nodeCenterX = nodeOffset.left - chartOffset.left + $node.outerWidth() / 2;

        var targetScrollLeft = nodeCenterX - $container.width() / 2;
        if (targetScrollLeft < 0) targetScrollLeft = 0;

        $container.scrollLeft(targetScrollLeft);
      }

      // Build the org chart
      var oc = $container.orgchart({
        data: orgData,
        nodeTitle: 'name',
        nodeContent: 'shortTitle',   // concise title on cards
        pan: true,
        zoom: true,
        draggable: true,
        direction: 't2b',
        visibleLevel: 3,
        exportButton: false,
        nodeTemplate: function(data) {
          // Group nodes: section headers
          if (data.isGroup) {
            var subtitle = data.title
              ? '<div class="org-node-title">' + data.title + '</div>'
              : '';
            return (
              '<div class="org-node-name">' + (data.name || '') + '</div>' +
              subtitle
            );
          }
          // People nodes: name + short title only
          var displayTitle = data.shortTitle || data.title || '';
          var titleLine = displayTitle
            ? '<div class="org-node-title">' + displayTitle + '</div>'
            : '';
          return (
            '<div class="org-node-name">' + (data.name || '') + '</div>' +
            titleLine
          );
        },
        createNode: function($node, data) {
          // Apply className from data for color-coding
          if (data.className) {
            $node.addClass(data.className);
          }

          // Classes for styling
          if (data.isGroup) {
            $node.addClass('group-node');
          } else if (data.isLeader) {
            $node.addClass('team-leader');
          }

          // Tooltip with full details (name — full title — org)
          var tooltipParts = [];
          if (data.name) tooltipParts.push(data.name);
          if (data.title) tooltipParts.push(data.title);
          if (data.org) tooltipParts.push(data.org);
          if (tooltipParts.length > 0) {
            $node.attr('title', tooltipParts.join(' — '));
          }

          // Focus style + recenter on click
          $node.on('click', function() {
            $('.orgchart .node').removeClass('focused');
            $(this).addClass('focused');
            centerOnNode($(this));
          });
        }
      });

      // Center global root after initial render
      centerChart();
      $(window).on('resize', centerChart);
    });
  </script>
</body>
</html>
'''

    # Inject the JSON safely
    html_with_data = html_template.replace("__ORG_DATA__", hierarchy_json)

    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html_with_data)

    print(f"[INFO] OrgChart HTML generated: {output_html}")

    # Open in default browser
    abs_path = os.path.abspath(output_html)
    webbrowser.open(f"file://{abs_path}")


# ===========================================
# MAIN PIPELINE
# ===========================================

if __name__ == "__main__":
    df_clean = load_and_clean_org_data()
    build_org_chart(df_clean, output_html=OUTPUT_HTML)
