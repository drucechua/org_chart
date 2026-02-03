from flask import Flask, render_template, request, Response
import pandas as pd
import json
from collections import defaultdict, deque

# -------------------------------------------
# FLASK SETUP
# -------------------------------------------
app = Flask(__name__)

# Column names
COL_ID = "Unique Identifier"
COL_NAME = "Name"
COL_REPORTS_TO = "Reports To"
COL_TITLE = "Line Detail 1"
COL_ORG = "Organization Name"


# ===========================================
# PART 1: CLEANING / NORMALIZATION
# ===========================================
def load_and_clean_org_data_from_file(file_storage, sheet_name: str = "Org Chart") -> pd.DataFrame:
    """
    Read uploaded Excel file and perform basic cleaning:
      - drop 'unfilled' and rows with 6+ digit ids
      - normalize name field
      - canonicalize IDs per Name
      - normalize Reports To field (position-based -> person canonical id)
      - remove self-referencing reports
    Raises RuntimeError with helpful text on failure to read the sheet.
    """
    try:
        df = pd.read_excel(file_storage, sheet_name=sheet_name)
    except Exception as e:
        raise RuntimeError(f"Could not read sheet '{sheet_name}': {e}")

    # Guard: ensure expected columns exist
    if COL_ID not in df.columns or COL_NAME not in df.columns:
        raise RuntimeError(f"Input file must contain columns: '{COL_ID}' and '{COL_NAME}'")

    # Remove unfilled rows & ids with 6+ digits
    uid_series = df[COL_ID].astype(str)
    df = df[
        ~uid_series.str.contains("unfilled", case=False, na=False)
        & ~uid_series.str.contains(r"\d{6,}", regex=True)
    ].copy()

    # Normalize name
    df[COL_NAME] = df[COL_NAME].astype(str).str.strip()

    # Canonical ID per Name (first seen)
    name_to_canonical_id = df.groupby(COL_NAME)[COL_ID].first().to_dict()

    # One row per person
    df_unique = df.drop_duplicates(subset=[COL_NAME], keep="first").copy()
    df_unique[COL_ID] = df_unique[COL_NAME].map(name_to_canonical_id)

    # Ensure Reports To column exists
    if COL_REPORTS_TO not in df_unique.columns:
        df_unique[COL_REPORTS_TO] = pd.NA

    # Normalize reporting lines of the form "<id>_<Name>"
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

    # Remove self-referencing reports
    mask_self = df_unique[COL_REPORTS_TO] == df_unique[COL_ID]
    df_unique.loc[mask_self, COL_REPORTS_TO] = pd.NA

    return df_unique


# ===========================================
# PART 2: ORG CHART GENERATION
# ===========================================
def is_null(x):
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


def build_org_chart_html(df):
    df = df.copy()

    df[COL_ID] = df[COL_ID].astype(str)
    df[COL_NAME] = df[COL_NAME].astype(str).str.strip()

    if COL_REPORTS_TO not in df.columns:
        df[COL_REPORTS_TO] = pd.NA

    for col in [COL_TITLE, COL_ORG]:
        if col not in df.columns:
            df[col] = ""

    # Build node records
    nodes = {}
    children_map = defaultdict(list)
    parent_map = {}

    for _, row in df.iterrows():
        uid = row[COL_ID]
        org_val = row.get(COL_ORG, "")

        # Title handling
        raw_title = row.get(COL_TITLE, "")
        full_title = raw_title.strip() if isinstance(raw_title, str) else ""
        leader_flag = is_leader_value(org_val)

        short_title = full_title or ""
        if isinstance(short_title, str) and "," in short_title:
            short_title = short_title.split(",")[0]
        short_title = short_title.strip()
        if len(short_title) > 40:
            short_title = short_title[:37].rstrip() + "…"

        title_lower = full_title.lower() if isinstance(full_title, str) else ""
        class_tokens = []
        if "trainee" in title_lower:
            class_tokens.append("role-trainee")
        else:
            class_tokens.append("role-leader" if leader_flag else "role-staff")

        nodes[uid] = {
            "id": uid,
            "name": row.get(COL_NAME, "") or "",
            "title": full_title,
            "shortTitle": short_title,
            "org": org_val or "",
            "children": [],
            "isLeader": leader_flag,
            "className": " ".join(class_tokens),
        }

    # Parent/child wiring (defensive: unknown manager -> treat as root)
    for _, row in df.iterrows():
        uid = row[COL_ID]
        manager_id = row[COL_REPORTS_TO]

        if is_null(manager_id):
            parent_map[uid] = None
        else:
            mid = str(manager_id)
            # If manager is known, attach; otherwise treat this person as root
            if mid in nodes:
                parent_map[uid] = mid
                children_map[mid].append(uid)
            else:
                parent_map[uid] = None

    # Attach children objects
    for parent_id, child_ids in children_map.items():
        if parent_id in nodes:
            nodes[parent_id]["children"] = [nodes[cid] for cid in child_ids if cid in nodes]

    # Find roots
    roots = [uid for uid, pid in parent_map.items() if pid is None]
    app.logger.info(f"Detected {len(roots)} root node(s): {roots}")

    if not roots:
        raise RuntimeError("No root nodes detected – cannot build org chart.")

    if len(roots) == 1:
        root_id = roots[0]
        root_node = nodes[root_id]
        existing_cls = nodes[root_id].get("className", "")
        nodes[root_id]["className"] = (existing_cls + " role-chro").strip()
    else:
        root_node = {
            "id": "VIRTUAL_ROOT",
            "name": "Organization",
            "title": "",
            "shortTitle": "",
            "org": "",
            "children": [nodes[rid] for rid in roots],
            "isLeader": True,
            "isGroup": True,
        }

    # Grouping under CHRO (virtual group nodes)
    group_ids = []
    if len(roots) == 1:
        original_children = root_node.get("children", [])

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
            "compact": True,
        }

        for child in original_children:
            title = (child.get("title") or "").lower()
            if "trainee" in title:
                group_trainees["children"].append(child)
            elif any(k in title for k in ["director", "head", "manager", "chief"]):
                group_leaders["children"].append(child)
            else:
                group_staff["children"].append(child)

        new_children = []
        for g in (group_leaders, group_staff, group_trainees):
            if g["children"]:
                new_children.append(g)
                group_ids.append(g["id"])

        if new_children:
            root_node["children"] = new_children

    # ---- Mark vertical / hybrid branches from level 4 onward ----
    def mark_vertical_branches(root, start_level=4, max_horizontal_children=3):
        """
        BFS walk:
         - assign node['orgLevel']
         - mark node['hybrid']=True only for person nodes (not isGroup)
           when they are at or below start_level and have more than
           max_horizontal_children *real person* children.

        We treat children with names starting with '(' (e.g. "(2) Managers")
        as aggregated placeholders and exclude them from the count.
        """
        queue = deque([(root, 1)])  # (node, level)
        hybrids = []

        while queue:
            node, level = queue.popleft()
            node["orgLevel"] = level

            children = node.get("children", []) or []

            # Filter out group nodes and aggregated placeholders like "(2) Managers"
            person_children = [
                c for c in children
                if not c.get("isGroup")
                and isinstance(c.get("name", ""), str)
                and not c.get("name", "").strip().startswith("(")
            ]

            # Mark hybrid only if this node is a person (not isGroup) AND it exceeds threshold
            if (level >= start_level
                and not node.get("isGroup")
                and len(person_children) > max_horizontal_children):
                node["hybrid"] = True
                hybrids.append((node.get("id"), node.get("name"), level, len(person_children)))

            # Enqueue all children to continue BFS — still assign orgLevel down the tree
            for child in children:
                queue.append((child, level + 1))

        # Debug prints only when Flask app is in debug mode
        if app.debug:
            if hybrids:
                app.logger.debug("[HYBRID] nodes marked (id, name, level, person_children_count):")
                for h in hybrids:
                    app.logger.debug(h)
            else:
                app.logger.debug("[HYBRID] no hybrids marked")

    mark_vertical_branches(root_node, start_level=4, max_horizontal_children=3)

    # ---- Collapse logic ----
    default_expanded_group_id = "GROUP_LEADERS" if "GROUP_LEADERS" in group_ids else None

    def apply_collapse_flags(node, is_root=False, expanded_group_id=None):
        children = node.get("children", [])
        if children:
            if is_root:
                node["collapsed"] = False
            elif node.get("isGroup"):
                node["collapsed"] = not (expanded_group_id and node.get("id") == expanded_group_id)
            else:
                node["collapsed"] = not node.get("isLeader", False)
            for c in children:
                apply_collapse_flags(c, expanded_group_id=expanded_group_id)

    apply_collapse_flags(root_node, is_root=True, expanded_group_id=default_expanded_group_id)

    # Serialize and inject into HTML
    hierarchy_json = json.dumps(root_node, indent=2)

    # HTML template with OrgChart integration. Kept simple and self-contained.
    html_template = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Org Chart</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <link rel="stylesheet"
        href="https://cdn.jsdelivr.net/npm/orgchart@3.8.0/dist/css/jquery.orgchart.min.css">
  <style>
    html, body { margin:0; padding:0; height:100%; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif; background:#f5f5f7; }
    #chart-container { width:100%; height:100vh; overflow:auto; background:linear-gradient(180deg,#f5f5f7 0%,#ffffff 40%); position:relative; }
    .hint-bar { position:absolute; top:10px; left:16px; background:rgba(17,24,39,0.78); color:#e5e7eb; padding:6px 12px; border-radius:999px; font-size:12px; z-index:10; }
    .orgchart { background:transparent !important; }
    .orgchart .nodes { align-items:flex-start; }
    .orgchart .node { border-radius:10px; box-shadow:0 2px 6px rgba(0,0,0,0.08); border:1px solid #d0d7de; background:#ffffff; padding:8px; box-sizing:border-box; }
    .orgchart .node:not(.group-node) { width:190px; min-height:72px; max-height:72px; display:flex; flex-direction:column; justify-content:center; text-align:center; }
    .org-node-name, .org-node-title { white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .org-node-name { font-weight:600; font-size:13px; margin-bottom:4px; color:#111827; }
    .org-node-title { font-size:11px; color:#4b5563; }
    .orgchart .node.team-leader { border-color:#2563eb; background:#eff6ff; }
    .orgchart .node.focused { border-color:#2563eb; box-shadow:0 0 0 2px rgba(37,99,235,0.18); }
    .orgchart .node.group-node { background:#111827; border-color:#111827; color:#f9fafb; box-shadow:0 4px 10px rgba(15,23,42,0.35); padding:10px 14px; }
    .group-node .org-node-name { font-size:14px; letter-spacing:0.04em; text-transform:uppercase; margin-bottom:6px; color:#f9fafb; }
    .group-node .org-node-title { font-size:11px; color:#e5e7eb; }
    .orgchart .node.role-chro { background:#020617; border-color:#020617; color:#f9fafb; }
    .orgchart .node.role-trainee:not(.group-node) { background:#f5f3ff; border-color:#a855f7; }
    .orgchart .level { padding-top:14px; padding-bottom:14px; }
    .orgchart .lines .topLine, .orgchart .lines .leftLine, .orgchart .lines .rightLine, .orgchart .lines .downLine { border-color:#e6e9ee; }

    /* Make the global root more prominent and readable */
    .orgchart .node.role-chro {
        background:#020617;
        border-color:#020617;
        color:#f9fafb;
        padding: 14px 18px;
        min-width: 260px;
        max-width: 420px;
    }
    .orgchart .node.role-chro .org-node-name {
        font-size: 14px;
        font-weight: 700;
        color: #f9fafb;
    }
    .orgchart .node.role-chro .org-node-title {
        font-size: 12px;
        color: #e5e7eb;
    }
  </style>
</head>
<body>
  <div class="hint-bar">Tip: Click a group or leader to expand their team.</div>
  <div id="chart-container"></div>
  <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/orgchart@3.8.0/dist/js/jquery.orgchart.min.js"></script>
  <script>
    var orgData = __ORG_DATA__;
    $(function() {
      var $container = $('#chart-container');
      function getChart() { return $container.find('.orgchart'); }
      function centerChart() {
        var $chart = getChart(); if (!$chart.length) return;
        var chartWidth = $chart.outerWidth(); var containerWidth = $container.width();
        if (chartWidth > containerWidth) { $container.scrollLeft((chartWidth - containerWidth)/2); } else { $container.scrollLeft(0); }
      }
      function centerOnNode($node) {
        var $chart = getChart(); if (!$chart.length || !$node.length) return;
        var chartOffset = $chart.offset(); var nodeOffset = $node.offset();
        var nodeCenterX = nodeOffset.left - chartOffset.left + $node.outerWidth()/2;
        var targetScrollLeft = nodeCenterX - $container.width()/2; if (targetScrollLeft < 0) targetScrollLeft = 0;
        $container.scrollLeft(targetScrollLeft);
      }
      var oc = $container.orgchart({
        data: orgData,
        nodeTitle: 'name',
        nodeContent: 'shortTitle',
        pan: true,
        zoom: true,
        draggable: false,
        direction: 't2b',
        visibleLevel: 3,
        verticalLevel: 4,
        nodeTemplate: function(data) {
          if (data.isGroup) {
            var subtitle = data.title ? '<div class="org-node-title">' + data.title + '</div>' : '';
            return '<div class="org-node-name">' + (data.name || '') + '</div>' + subtitle;
          }
          var displayTitle = data.shortTitle || data.title || '';
          var titleLine = displayTitle ? '<div class="org-node-title">' + displayTitle + '</div>' : '';
          return '<div class="org-node-name">' + (data.name || '') + '</div>' + titleLine;
        },
        createNode: function($node, data) {
          if (data.className) { $node.addClass(data.className); }
          if (data.isGroup) { $node.addClass('group-node'); } else if (data.isLeader) { $node.addClass('team-leader'); }
          var tooltipParts = []; if (data.name) tooltipParts.push(data.name); if (data.title) tooltipParts.push(data.title); if (data.org) tooltipParts.push(data.org);
          if (tooltipParts.length > 0) { $node.attr('title', tooltipParts.join(' — ')); }
          $node.on('click', function() {
            $('.orgchart .node').removeClass('focused');
            $(this).addClass('focused');
            centerOnNode($(this));
          });
        }
      });
      // center globally, then focus the root for readability
      centerChart();
      setTimeout(function() {
        var $root = $container.find('.orgchart .node.role-chro').first();
        if (!$root.length) $root = $container.find('.orgchart .node').first();
        if ($root.length) {
          centerOnNode($root);
          $('.orgchart .node').removeClass('focused');
          $root.addClass('focused');
        }
      }, 220);
      $(window).on('resize', centerChart);
    });
  </script>
</body>
</html>
"""

    html_with_data = html_template.replace("__ORG_DATA__", hierarchy_json)
    return html_with_data


# ===========================================
# FLASK ROUTES
# ===========================================
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            return "No file uploaded", 400

        try:
            df_clean = load_and_clean_org_data_from_file(file)
            html = build_org_chart_html(df_clean)
            return Response(html, mimetype="text/html")
        except Exception as e:
            # return a helpful error for users; in prod you'd log it as well
            return f"Error processing file: {e}", 500

    # GET request -> show upload page
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
