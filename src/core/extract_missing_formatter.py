"""
Extract missing data formatter module.

Extracts rows from an input Excel file that correspond to styles
in the output file where the "AI Generated Image Flag" column is missing or empty.
"""

import pandas as pd
from pathlib import Path
from dataclasses import dataclass
from typing import Optional


@dataclass
class ExtractMissingResult:
    """Result of an extract missing data operation."""
    success: bool
    output_path: Optional[Path] = None
    rows_extracted: int = 0
    types_rows: int = 0
    missing_count: int = 0
    error_message: Optional[str] = None


class ExcelValidationError(Exception):
    """Custom exception for Excel validation errors."""
    pass


def validate_excel_file(file_path: Path, file_name: str) -> None:
    """
    Validate that an Excel file is readable and not empty.

    Args:
        file_path: Path to the Excel file
        file_name: Name of the file for error messages

    Raises:
        ExcelValidationError: If file is invalid or empty
    """
    try:
        # Try reading the file
        df = pd.read_excel(file_path)

        if df is None:
            raise ExcelValidationError(f"{file_name}: Could not read Excel file (file may be corrupted)")

        if df.empty:
            raise ExcelValidationError(f"{file_name}: Excel file contains no data")

        # Check for valid columns
        if len(df.columns) == 0:
            raise ExcelValidationError(f"{file_name}: Excel file has no columns")

    except pd.errors.EmptyDataError:
        raise ExcelValidationError(f"{file_name}: Excel file is empty or contains no parseable data")
    except pd.errors.ParserError as e:
        raise ExcelValidationError(f"{file_name}: Could not parse Excel file: {e}")
    except Exception as e:
        if isinstance(e, ExcelValidationError):
            raise
        raise ExcelValidationError(f"{file_name}: Error reading file: {e}")


def extract_rows_with_missing_ai_flag(
    input_file: Path,
    output_file: Path,
    result_file: Path,
    values_sheet: str = "Values",
    types_sheet: str = "Types",
    style_id_col: str = "styleId",
    flag_col: str = "AI Generated Image Flag"
) -> ExtractMissingResult:
    """
    Extract rows from input file where AI flag is missing in output file.

    This function:
    1. Reads the output file to find styleIds where AI Generated Image Flag is missing/empty
    2. Reads the input Excel file's "Values" and "Types" sheets
    3. Filters the "Values" sheet to include only rows where styleId has missing AI flag
    4. Writes both filtered "Values" and original "Types" sheets to a new Excel file

    Args:
        input_file: Path to the input Excel file (with Values and Types sheets)
        output_file: Path to the output Excel file (with AI Generated Image Flag column)
        result_file: Path for the result Excel file
        values_sheet: Name of the Values sheet (default: "Values")
        types_sheet: Name of the Types sheet (default: "Types")
        style_id_col: Name of the style ID column (default: "styleId")
        flag_col: Name of the AI flag column (default: "AI Generated Image Flag")

    Returns:
        ExtractMissingResult with operation details
    """
    try:
        # Validate input files
        validate_excel_file(input_file, "Input file")
        validate_excel_file(output_file, "Output file")

        # Read output file to find styleIds with missing AI flag
        output_df = pd.read_excel(output_file)

        # Check if required columns exist
        if flag_col not in output_df.columns:
            return ExtractMissingResult(
                success=False,
                error_message=f"Output file must contain '{flag_col}' column"
            )

        if style_id_col not in output_df.columns:
            return ExtractMissingResult(
                success=False,
                error_message=f"Output file must contain '{style_id_col}' column"
            )

        # Find styleIds where AI Generated Image Flag is missing or empty
        missing_flag_df = output_df[
            output_df[flag_col].isna() |
            (output_df[flag_col].astype(str).str.strip() == "")
        ]

        missing_style_ids = set(missing_flag_df[style_id_col].dropna())

        if len(missing_style_ids) == 0:
            return ExtractMissingResult(
                success=True,
                output_path=result_file,
                rows_extracted=0,
                types_rows=0,
                missing_count=0,
                error_message=None
            )

        # Read input sheets
        try:
            input_values_df = pd.read_excel(input_file, sheet_name=values_sheet)
        except ValueError:
            return ExtractMissingResult(
                success=False,
                error_message=f"Input file must contain a '{values_sheet}' sheet"
            )

        try:
            input_types_df = pd.read_excel(input_file, sheet_name=types_sheet)
        except ValueError:
            return ExtractMissingResult(
                success=False,
                error_message=f"Input file must contain a '{types_sheet}' sheet"
            )

        # Check if styleId column exists in Values sheet
        if style_id_col not in input_values_df.columns:
            return ExtractMissingResult(
                success=False,
                error_message=f"Values sheet must contain '{style_id_col}' column"
            )

        # Filter Values sheet
        filtered_values_df = input_values_df[
            input_values_df[style_id_col].isin(missing_style_ids)
        ]

        # Write both sheets to output Excel
        with pd.ExcelWriter(result_file, engine="openpyxl") as writer:
            input_types_df.to_excel(writer, sheet_name="Types", index=False)
            filtered_values_df.to_excel(writer, sheet_name="Values", index=False)

        return ExtractMissingResult(
            success=True,
            output_path=result_file,
            rows_extracted=len(filtered_values_df),
            types_rows=len(input_types_df),
            missing_count=len(missing_style_ids)
        )

    except FileNotFoundError as e:
        return ExtractMissingResult(
            success=False,
            error_message=f"File not found: {e}"
        )
    except ExcelValidationError as e:
        return ExtractMissingResult(
            success=False,
            error_message=str(e)
        )
    except Exception as e:
        return ExtractMissingResult(
            success=False,
            error_message=f"Error processing file: {e}"
        )
