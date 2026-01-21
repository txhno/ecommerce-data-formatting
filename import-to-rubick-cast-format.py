#!/usr/bin/env python3
"""
merge_sizechart_productdetails_with_types_values.py

Final version:
• Aggregates size chart rows per styleId
• Merges ALL tabs into ONE dataset
• Canonical column mapping (case + whitespace insensitive)
• Union of all columns across tabs
• Outputs Excel with:
    - Types sheet (exact CAST format)
    - Values sheet (merged data)
"""

from pathlib import Path
import pandas as pd
import re
import warnings
from tqdm import tqdm

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------- CONFIG ----------------
SIZE_CHART_PATH = Path("sku.xlsx")
PRODUCT_DETAILS_PATH = Path("style.xlsx")
OUTPUT_PATH = Path("Batch_Merged_With_Types_Values.xlsx")

EXCLUDE_SHEETS = {"masterdata"}

# ---------------- HELPERS ----------------
STYLE_PATTERNS = [
    re.compile(r"^style[_\s\-]?id$", re.I),
    re.compile(r"^sku$", re.I),
    re.compile(r"styleid", re.I),
]

IMAGE_PATTERNS = [
    re.compile(p, re.I) for p in [
        "image", "img", "url", "cdn"
    ]
]

def normalize_col(col: str) -> str:
    return re.sub(r"\s+", "", str(col).lower())

def find_style_col(columns):
    for c in columns:
        for p in STYLE_PATTERNS:
            if p.search(str(c)):
                return c
    return None

def find_brand_size_start(columns):
    for c in columns:
        s = str(c).lower()
        if "brand" in s and "size" in s:
            return c
    for c in columns:
        if "size" in str(c).lower():
            return c
    return None

def aggregate_list(series):
    vals = (
        series.dropna()
        .astype(str)
        .str.strip()
        .replace({"": None, "nan": None, "none": None})
        .dropna()
        .unique()
    )
    return ",".join(vals) if len(vals) else None

def infer_type(col_name: str) -> str:
    for p in IMAGE_PATTERNS:
        if p.search(col_name):
            return "image"
    return "string"

# ---------------- LOAD FILES ----------------
size_xl = pd.ExcelFile(SIZE_CHART_PATH, engine="openpyxl")
prod_xl = pd.ExcelFile(PRODUCT_DETAILS_PATH, engine="openpyxl")

size_sheets = [s for s in size_xl.sheet_names if s not in EXCLUDE_SHEETS]
prod_sheets = set(prod_xl.sheet_names)

canonical_cols = {}
ordered_columns = []
final_dfs = []

# ---------------- PROCESS SHEETS ----------------
for sheet in tqdm(size_sheets, desc="Processing Sheets", ncols=90):

    size_df = size_xl.parse(sheet, dtype=str)
    prod_df = (
        prod_xl.parse(sheet, dtype=str)
        if sheet in prod_sheets
        else pd.DataFrame()
    )

    size_df.columns = [str(c) for c in size_df.columns]
    prod_df.columns = [str(c) for c in prod_df.columns]

    style_size = find_style_col(size_df.columns)
    style_prod = find_style_col(prod_df.columns) if not prod_df.empty else None

    if style_size is None:
        continue

    if prod_df.empty:
        prod_df = pd.DataFrame({style_size: size_df[style_size].unique()})
        style_prod = style_size

    if style_prod is None:
        style_prod = style_size
        prod_df[style_prod] = prod_df.iloc[:, 0]

    # Register product columns
    for col in prod_df.columns:
        key = normalize_col(col)
        if key not in canonical_cols:
            canonical_cols[key] = col
            ordered_columns.append(col)

    prod_df.rename(columns=lambda c: canonical_cols[normalize_col(c)], inplace=True)

    # -------- SIZE AGGREGATION --------
    brand_col = find_brand_size_start(size_df.columns)

    if brand_col:
        size_cols = size_df.columns[size_df.columns.get_loc(brand_col):]

        melted = size_df.melt(
            id_vars=[style_size],
            value_vars=size_cols,
            var_name="col",
            value_name="val",
        )

        aggregated = (
            melted.groupby([style_size, "col"])["val"]
            .apply(aggregate_list)
            .reset_index()
        )

        pivot = aggregated.pivot(
            index=style_size, columns="col", values="val"
        ).reset_index()

        pivot.rename(columns={style_size: style_prod}, inplace=True)

        prod_df = prod_df.merge(pivot, on=style_prod, how="outer")

        for col in pivot.columns:
            if col == style_prod:
                continue
            key = normalize_col(col)
            if key not in canonical_cols:
                canonical_cols[key] = col
                ordered_columns.append(col)

    final_dfs.append(prod_df)

# ---------------- FINAL VALUES DF ----------------
final_df = pd.concat(final_dfs, ignore_index=True)

for col in ordered_columns:
    if col not in final_df.columns:
        final_df[col] = None

final_df = final_df[ordered_columns]
final_df = final_df.where(pd.notnull(final_df), "")

# ---------------- BUILD TYPES SHEET ----------------
types_columns = ["Column1", "Column2"] + ordered_columns
types_df = pd.DataFrame(columns=types_columns)

# Row 2: column names repeated
for col in ordered_columns:
    types_df.loc[0, col] = col

# Row 3: mandatory
for col in ordered_columns:
    types_df.loc[1, col] = "mandatory"

# Row 4: types
for col in ordered_columns:
    types_df.loc[2, col] = infer_type(col)

# ---------------- WRITE EXCEL ----------------
with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter") as writer:
    types_df.to_excel(writer, sheet_name="Types", index=False)
    final_df.to_excel(writer, sheet_name="Values", index=False)

print("\nProcess complete.")
print(f"Output file: {OUTPUT_PATH}")
print(f"Values rows : {len(final_df)}")
print(f"Columns     : {len(final_df.columns)}")
