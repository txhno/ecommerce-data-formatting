"""
Import formatter module for merging size chart and product details.

Aggregates size chart rows per styleId, merges all tabs into one dataset,
and outputs Excel with Types and Values sheets in CAST format.
"""

import re
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

import pandas as pd


@dataclass
class ImportResult:
    """Result of an import merging operation."""
    success: bool
    output_path: Optional[Path] = None
    rows_processed: int = 0
    columns_count: int = 0
    sheets_processed: int = 0
    error_message: Optional[str] = None


class ImportValidationError(Exception):
    """Custom exception for import validation errors."""
    pass


def normalize_col(col: str) -> str:
    """Normalize column name for canonical mapping."""
    return re.sub(r"\s+", "", str(col).lower())


def find_style_column(columns, sheet_name: str = "sheet") -> tuple[Optional[str], Optional[str]]:
    """
    Find the style ID column in a list of columns.

    Returns:
        Tuple of (found_column, suggestion) or (None, None) if not found
    """
    patterns = [
        (re.compile(r"^style[_\s\-]?id$", re.I), "style_id"),
        (re.compile(r"^sku$", re.I), "SKU"),
        (re.compile(r"styleid", re.I), "styleId"),
    ]
    for c in columns:
        for p, suggestion in patterns:
            if p.search(str(c)):
                return c, None
    # Return first column as suggestion if no style column found
    if len(columns) > 0:
        return None, str(columns[0])
    return None, None


def find_brand_size_column(columns) -> Optional[str]:
    """Find the brand size column (starts brand size columns)."""
    for c in columns:
        s = str(c).lower()
        if "brand" in s and "size" in s:
            return c
    for c in columns:
        if "size" in str(c).lower():
            return c
    return None


def validate_excel_file(file_path: Path, file_name: str) -> None:
    """
    Validate that an Excel file is readable and has valid structure.

    Args:
        file_path: Path to the Excel file
        file_name: Name of the file for error messages

    Raises:
        ImportValidationError: If file is invalid or empty
    """
    try:
        xl = pd.ExcelFile(file_path, engine="openpyxl")

        if len(xl.sheet_names) == 0:
            raise ImportValidationError(f"{file_name}: Excel file has no sheets")

        # Check first sheet has data
        df = xl.parse(xl.sheet_names[0], dtype=str)
        if df.empty and len(xl.sheet_names) == 1:
            raise ImportValidationError(f"{file_name}: First sheet is empty")
        if len(df.columns) == 0:
            raise ImportValidationError(f"{file_name}: Sheet has no columns")

    except ImportValidationError:
        raise
    except pd.errors.EmptyDataError:
        raise ImportValidationError(f"{file_name}: Excel file is empty")
    except pd.errors.ParserError as e:
        raise ImportValidationError(f"{file_name}: Could not parse Excel file: {e}")
    except Exception as e:
        if isinstance(e, ImportValidationError):
            raise
        raise ImportValidationError(f"{file_name}: Error reading file: {e}")


def aggregate_list(series) -> Optional[str]:
    """Aggregate multiple values into a comma-separated string."""
    vals = (
        series.dropna()
        .astype(str)
        .str.strip()
        .replace({"": None, "nan": None, "none": None})
        .dropna()
        .unique()
    )
    return ",".join(vals) if len(vals) else None


def infer_column_type(col_name: str) -> str:
    """Infer the data type for a column based on its name."""
    image_patterns = ["image", "img", "url", "cdn"]
    for pattern in image_patterns:
        if re.search(pattern, col_name, re.I):
            return "image"
    return "string"


def merge_sizechart_productdetails(
    size_chart_path: Path,
    product_details_path: Path,
    output_path: Path,
    exclude_sheets: Optional[list[str]] = None
) -> ImportResult:
    """
    Merge size chart and product details into CAST format.

    This function:
    1. Loads both input Excel files
    2. Aggregates size chart rows per styleId
    3. Merges all sheets into one dataset
    4. Creates Types sheet with column metadata
    5. Creates Values sheet with merged data

    Args:
        size_chart_path: Path to the size chart Excel file
        product_details_path: Path to the product details Excel file
        output_path: Path for the merged output Excel file
        exclude_sheets: List of sheet names to exclude

    Returns:
        ImportResult with operation details
    """
    try:
        exclude_sheets = exclude_sheets or ["masterdata"]

        # Validate input files
        validate_excel_file(size_chart_path, "Size chart file")
        validate_excel_file(product_details_path, "Product details file")

        size_xl = pd.ExcelFile(size_chart_path, engine="openpyxl")
        prod_xl = pd.ExcelFile(product_details_path, engine="openpyxl")

        size_sheets = [s for s in size_xl.sheet_names if s not in exclude_sheets]
        prod_sheets = set(prod_xl.sheet_names)

        if not size_sheets:
            return ImportResult(
                success=False,
                error_message="No valid sheets found in size chart after exclusions"
            )

        canonical_cols = {}
        ordered_columns = []
        final_dfs = []

        # Track if any sheet had valid data
        any_valid_data = False

        # Process each sheet
        for sheet in size_sheets:
            size_df = size_xl.parse(sheet, dtype=str)
            prod_df = (
                prod_xl.parse(sheet, dtype=str)
                if sheet in prod_sheets
                else pd.DataFrame()
            )

            size_df.columns = [str(c) for c in size_df.columns]
            prod_df.columns = [str(c) for c in prod_df.columns]

            style_size, suggestion = find_style_column(size_df.columns, sheet)

            if style_size is None:
                # Only error on first sheet - subsequent sheets might be intentionally empty
                if sheet == size_sheets[0]:
                    columns_list = ", ".join(size_df.columns[:5].tolist())
                    if len(size_df.columns) > 5:
                        columns_list += ", ..."
                    return ImportResult(
                        success=False,
                        error_message=f"Could not find style ID column in '{sheet}'. "
                                      f"Expected column like 'style_id', 'SKU', or 'styleId'. "
                                      f"Found: [{columns_list}]. "
                                      f"Rename your identifier column to match one of these patterns."
                    )
                continue

            any_valid_data = True

            if prod_df.empty:
                prod_df = pd.DataFrame({style_size: size_df[style_size].unique()})
                style_prod = style_size

            if prod_df.empty:
                prod_df = pd.DataFrame({style_size: size_df[style_size].dropna().unique()})
                style_prod = style_size
            else:
                style_prod = find_style_column(prod_df.columns)[0] if not prod_df.empty else style_size

            if style_prod is None:
                style_prod = style_size
                if not prod_df.empty:
                    prod_df[style_prod] = prod_df.iloc[:, 0]

            # Register product columns
            for col in prod_df.columns:
                key = normalize_col(col)
                if key not in canonical_cols:
                    canonical_cols[key] = col
                    ordered_columns.append(col)

            if not prod_df.empty:
                prod_df.rename(columns=lambda c: canonical_cols[normalize_col(c)], inplace=True)

            # Size aggregation
            brand_col = find_brand_size_column(size_df.columns)

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

        if not any_valid_data:
            return ImportResult(
                success=False,
                error_message="No valid data found in any sheet. Check that files contain style ID columns."
            )

        # Final values dataframe
        final_df = pd.concat(final_dfs, ignore_index=True)

        if final_df.empty:
            return ImportResult(
                success=False,
                error_message="No data produced after merging. Check input file formats."
            )

        for col in ordered_columns:
            if col not in final_df.columns:
                final_df[col] = None

        final_df = final_df[ordered_columns]
        final_df = final_df.where(pd.notnull(final_df), "")

        # Build Types sheet
        types_columns = ["Column1", "Column2"] + ordered_columns
        types_df = pd.DataFrame(columns=types_columns)

        for col in ordered_columns:
            types_df.loc[0, col] = col
        for col in ordered_columns:
            types_df.loc[1, col] = "mandatory"
        for col in ordered_columns:
            types_df.loc[2, col] = infer_column_type(col)

        # Write Excel output
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            types_df.to_excel(writer, sheet_name="Types", index=False)
            final_df.to_excel(writer, sheet_name="Values", index=False)

        return ImportResult(
            success=True,
            output_path=output_path,
            rows_processed=len(final_df),
            columns_count=len(final_df.columns),
            sheets_processed=len(size_sheets)
        )

    except FileNotFoundError as e:
        return ImportResult(
            success=False,
            error_message=f"File not found: {e}"
        )
    except ImportValidationError as e:
        return ImportResult(
            success=False,
            error_message=str(e)
        )
    except Exception as e:
        return ImportResult(
            success=False,
            error_message=f"Error processing files: {e}"
        )
