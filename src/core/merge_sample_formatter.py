"""
Merge sample output formatter module.

Merges updates from a sample output file into an existing output file
by overwriting rows based on styleId.
"""

import pandas as pd
from pathlib import Path
from dataclasses import dataclass
from typing import Optional


@dataclass
class MergeSampleResult:
    """Result of a merge sample output operation."""
    success: bool
    output_path: Optional[Path] = None
    rows_updated: int = 0
    total_rows: int = 0
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


def merge_sample_output(
    output_file: Path,
    sample_file: Path,
    result_file: Path,
    style_id_col: str = "styleId"
) -> MergeSampleResult:
    """
    Merge sample output updates into main output file.

    This function:
    1. Reads both output and sample Excel files
    2. Verifies styleId column exists in both files
    3. Uses styleId as index for safe row matching
    4. Overwrites rows in output file with corresponding rows from sample file
    5. Only updates columns that exist in both files (prevents schema mismatch)
    6. Outputs final merged file

    Args:
        output_file: Path to the main output file to be updated
        sample_file: Path to the sample file containing updated rows
        result_file: Path for the merged result file
        style_id_col: Name of the style ID column (default: "styleId")

    Returns:
        MergeSampleResult with operation details
    """
    try:
        # Validate input files
        validate_excel_file(output_file, "Output file")
        validate_excel_file(sample_file, "Sample file")

        # Read files
        output_df = pd.read_excel(output_file)
        sample_df = pd.read_excel(sample_file)

        # Ensure styleId exists in both files
        if style_id_col not in output_df.columns:
            return MergeSampleResult(
                success=False,
                error_message=f"Output file must contain '{style_id_col}' column"
            )

        if style_id_col not in sample_df.columns:
            return MergeSampleResult(
                success=False,
                error_message=f"Sample file must contain '{style_id_col}' column"
            )

        # Set index to styleId for safe overwrite
        output_df_indexed = output_df.set_index(style_id_col)
        sample_df_indexed = sample_df.set_index(style_id_col)

        # Keep only columns common to both (prevents schema mismatch)
        common_columns = output_df_indexed.columns.intersection(sample_df_indexed.columns)

        # Track number of rows to be updated
        styles_to_update = sample_df_indexed.index.intersection(output_df_indexed.index)
        rows_updated = len(styles_to_update)

        # Overwrite rows in output with sample rows
        for style_id in sample_df_indexed.index:
            if style_id in output_df_indexed.index:
                output_df_indexed.loc[style_id, common_columns] = sample_df_indexed.loc[style_id, common_columns]

        # Reset index back
        output_df_final = output_df_indexed.reset_index()

        # Write final output
        output_df_final.to_excel(result_file, index=False)

        return MergeSampleResult(
            success=True,
            output_path=result_file,
            rows_updated=rows_updated,
            total_rows=len(output_df_final)
        )

    except FileNotFoundError as e:
        return MergeSampleResult(
            success=False,
            error_message=f"File not found: {e}"
        )
    except ExcelValidationError as e:
        return MergeSampleResult(
            success=False,
            error_message=str(e)
        )
    except Exception as e:
        return MergeSampleResult(
            success=False,
            error_message=f"Error processing file: {e}"
        )
