from flask import Flask, render_template, request
import math
import pandas as pd
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


def _json_safe_value(v):
    """Convert a value to something JSON-serializable (no NaN/pd.NA)."""
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (dict, list)):
        return v
    return v


def _sanitize_org_node(node):
    """Recursively sanitize org chart node so it is JSON-serializable."""
    if not isinstance(node, dict):
        return node
    out = {}
    string_keys = ("id", "name", "title", "shortTitle", "org", "className")
    bool_keys = ("isLeader", "isGroup", "collapsed", "compact", "hybrid")
    for k, val in node.items():
        if k == "children":
            out[k] = [_sanitize_org_node(c) for c in (val or [])]
        else:
            out[k] = _json_safe_value(val)
            if out[k] is None:
                if k in string_keys:
                    out[k] = ""
                elif k in bool_keys:
                    out[k] = False
    return out


def build_org_chart_data(df):
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
            "name": "LEADERSHIP",
            "title": "",
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
            "title": "",
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
            "title": "",
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

    return root_node


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
            org_data = build_org_chart_data(df_clean)
            org_data = _sanitize_org_node(org_data)
            return render_template("org_chart.html", org_data=org_data)
        except Exception as e:
            # return a helpful error for users; in prod you'd log it as well
            return f"Error processing file: {e}", 500

    # GET request -> show upload page
    return render_template("index.html")


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=True, host="0.0.0.0", port=port)
