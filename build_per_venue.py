import pandas as pd, re

SRC = "/root/.claude/uploads/25fd9b56-1752-5edc-8721-3932cd50f9d6/70eb99d7-20260523_naming_rights_database_1.xlsx"
ENR = "/home/user/NYU_MSBAi/first_50_venues_enriched.xlsx"

# Read the already-enriched full sheet (all original cols + added cols, 50 rows)
full = pd.read_excel(ENR, sheet_name="Full_Enriched")
full = full.head(50)

def clean_sheet_name(idx, name):
    name = re.sub(r'[:\\/?*\[\]]', '-', str(name))
    prefix = f"{idx:02d}_"
    return (prefix + name)[:31]

out = "/home/user/NYU_MSBAi/first_50_venues_by_venue.xlsx"
seen = set()
with pd.ExcelWriter(out, engine="openpyxl") as w:
    # Index sheet
    idx_df = full[["Record ID","Venue Name","City","State / Prefecture","Country"]].copy()
    idx_df.insert(0, "Sheet #", range(1, len(full)+1))
    idx_df.to_excel(w, sheet_name="Index", index=False)

    for i, (_, row) in enumerate(full.iterrows(), 1):
        sn = clean_sheet_name(i, row["Venue Name"])
        # ensure uniqueness
        base = sn
        k = 1
        while sn in seen:
            sn = (base[:29] + f"_{k}")[:31]; k += 1
        seen.add(sn)
        # vertical field | value, dropping empty values
        recs = [(col, row[col]) for col in full.columns
                if pd.notna(row[col]) and str(row[col]).strip() != ""]
        vdf = pd.DataFrame(recs, columns=["Field", "Value"])
        vdf.to_excel(w, sheet_name=sn, index=False)

print("Wrote", out)
# verify
xl = pd.ExcelFile(out)
print("Sheets:", len(xl.sheet_names))
print(xl.sheet_names[:6], "...")
