import pandas as pd
import re

# -------------------------------------------
# STEP 1 — LOAD ORIGINAL FILE
# -------------------------------------------
path = "Office of Human Resources  (AlNuaimi, Rashed).xlsx"
df = pd.read_excel(path, sheet_name="Org Chart")

# -------------------------------------------
# STEP 2 — REMOVE UNFILLED POSITIONS
# And remove rows where Unique Identifier contains more than 5 digits
# -------------------------------------------

# Convert column to string once
uid = df["Unique Identifier"].astype(str)

df = df[
    ~uid.str.contains("unfilled", case=False, na=False)  # remove unfilled rows
    & ~uid.str.contains(r"\d{6,}", regex=True)           # remove IDs with 6+ digits
]

# -------------------------------------------
# STEP 3 — STANDARDIZE NAME FIELD
# (Fix: use .str.strip())
# -------------------------------------------
df["Name"] = df["Name"].astype(str).str.strip()


# -------------------------------------------
# STEP 4 — BUILD CANONICAL PERSON-LEVEL ID
# Choose the first Unique Identifier for each Name
# -------------------------------------------
name_to_canonical_id = df.groupby("Name")["Unique Identifier"].first().to_dict()

# -------------------------------------------
# STEP 5 — DROP DUPLICATES (ONE ROW PER PERSON)
# -------------------------------------------
df_unique = df.drop_duplicates(subset=["Name"], keep="first").copy()

# Replace their Unique Identifier with canonical version
df_unique["Unique Identifier"] = df_unique["Name"].map(name_to_canonical_id)

# -------------------------------------------
# STEP 6 — NORMALIZE REPORTING LINES (NAME-BASED)
# Convert old position-based IDs to canonical person IDs
# -------------------------------------------
def extract_name_from_id(uid):
    if pd.isna(uid):
        return None
    uid = str(uid)
    parts = uid.split("_", 1)
    if len(parts) == 2:
        name = parts[1]
        return name.replace("_", " ")
    return None

def normalize_reports_to(uid):
    if pd.isna(uid):
        return None
    name = extract_name_from_id(uid)
    if name in name_to_canonical_id:
        return name_to_canonical_id[name]
    return None

df_unique["Reports To"] = df_unique["Reports To"].apply(normalize_reports_to)

# -------------------------------------------
# STEP 7 — REMOVE SELF-REFERENCING REPORTS
# -------------------------------------------
df_unique.loc[df_unique["Reports To"] == df_unique["Unique Identifier"], "Reports To"] = pd.NA

# -------------------------------------------
# STEP 8 — SAVE IDEAL FINAL OUTPUT FILE
# -------------------------------------------
output_path = "ideal_final_output.xlsx"
df_unique.to_excel(output_path, index=False)

print("Transformation complete!")
print(f"Saved as: {output_path}")
