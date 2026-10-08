import pandas as pd
from pathlib import Path

# --- Paths ---
input_file = Path(r"F:\Work\Projects\Excel_Data_Cleaning\Data\Cola.xlsx")
output_folder = Path(r"F:\Work\Projects\Excel_Data_Cleaning\Output")
output_folder.mkdir(parents=True, exist_ok=True)

# 1. Load raw without headers
raw = pd.read_excel(input_file, sheet_name="COCA COLA CO", header=None, engine="openpyxl")

# 2. Identify the three statement sections and their year headers
section_names = ["Profit & Loss statement", "Balance Sheet", "Cash Flow statement"]
sections = []
current_section = None

for idx, row in raw.iterrows():
    col_a = row.get(0)  # column A: section names
    col_b = row.get(1)  # column B: line items + "in million USD"

    if isinstance(col_a, str):
        col_a = col_a.strip()
        if col_a in section_names:
            current_section = {"name": col_a, "start": idx}
            sections.append(current_section)

    if isinstance(col_b, str) and col_b.strip() == "in million USD" and current_section is not None:
        years = [row.get(c) for c in range(2, raw.shape[1])]
        current_section["years"] = years
        current_section["data_start"] = idx + 1

# 3. Determine where each section ends
for i, sec in enumerate(sections):
    sec["end"] = sections[i + 1]["start"] if i + 1 < len(sections) else len(raw)

# 4. Extract every line item into a tidy long format
records = []
for sec in sections:
    if "years" not in sec:
        continue
    seen = {}
    for i in range(sec["data_start"], sec["end"]):
        row = raw.iloc[i]
        line_item = row.get(1)
        if pd.isna(line_item):
            continue
        line_item = str(line_item).strip()
        if line_item == "" or line_item in section_names or line_item == "in million USD":
            continue

        if line_item in seen:
            seen[line_item] += 1
            unique_item = f"{line_item} ({seen[line_item]})"
        else:
            seen[line_item] = 1
            unique_item = line_item

        for j, year in enumerate(sec["years"]):
            val = row.get(2 + j)
            if pd.isna(val):
                continue
            records.append({
                "statement": sec["name"],
                "line_item": unique_item,
                "fiscal_year": year,
                "value": val
            })

tidy = pd.DataFrame(records)

# --- Sanity check ---
print(f"Sections found: {[s['name'] for s in sections]}")
print(f"Rows extracted: {len(tidy)}")
if tidy.empty:
    raise SystemExit("No data extracted — check section detection.")

# 5. Clean year and value columns
tidy["fiscal_year"] = tidy["fiscal_year"].astype(str).str.extract(r"(\d+)")[0]
tidy["fiscal_year"] = pd.to_numeric(tidy["fiscal_year"], errors="coerce")
tidy.loc[tidy["fiscal_year"] < 100, "fiscal_year"] += 2000
tidy["fiscal_year"] = tidy["fiscal_year"].astype("Int64")

tidy["value"] = tidy["value"].astype(str).str.replace(",", "", regex=False)
tidy["value"] = pd.to_numeric(tidy["value"], errors="coerce")

# 6. Pivot each statement wide and save
for sec_name in tidy["statement"].unique():
    df_wide = tidy[tidy["statement"] == sec_name].pivot_table(
        index="line_item",
        columns="fiscal_year",
        values="value",
        aggfunc="first"
    ).sort_index(axis=1)

    print(f"\n=== {sec_name} ===")
    print(df_wide.head())

    safe_name = sec_name.replace(" ", "_").replace("&", "and")
    df_wide.to_excel(output_folder / f"{safe_name}_clean.xlsx")

# Save tidy long format
tidy.to_excel(output_folder / "Cola_tidy.xlsx", index=False)
print(f"\nDone. Files saved to: {output_folder}")