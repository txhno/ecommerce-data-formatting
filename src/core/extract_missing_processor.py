"""
Extract missing data processor module for Streamlit integration.

Provides a clean interface for the web UI to process extract missing data operations
without dealing with file handling details.
"""

import tempfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional

from src.utils.logging import get_logger
from .extract_missing_formatter import extract_rows_with_missing_ai_flag, ExtractMissingResult

logger = get_logger(__name__)


@dataclass
class ExtractMissingProcessorResult:
    """Result of an extract missing data processing operation for Streamlit."""
    success: bool
    data: Optional[BytesIO] = None
    filename: Optional[str] = None
    rows_extracted: int = 0
    types_rows: int = 0
    missing_count: int = 0
    error_message: Optional[str] = None


# Maximum file size in bytes (50MB)
MAX_FILE_SIZE = 50 * 1024 * 1024


def validate_file_size(data: bytes, filename: str, max_size: int = MAX_FILE_SIZE) -> Optional[str]:
    """
    Validate file size is within limits.

    Args:
        data: File data in bytes
        filename: Filename for error messages
        max_size: Maximum allowed size in bytes

    Returns:
        Error message if invalid, None if valid
    """
    if len(data) > max_size:
        size_mb = len(data) / (1024 * 1024)
        max_mb = max_size / (1024 * 1024)
        return f"{filename} is {size_mb:.1f}MB, exceeds maximum size of {max_mb:.0f}MB"
    return None


def process_extract_missing(
    input_file_data: bytes,
    input_filename: str,
    output_file_data: bytes,
    output_filename: str,
    result_filename: str = "styles_with_missing_ai_flag.xlsx",
    values_sheet: str = "Values",
    types_sheet: str = "Types",
    style_id_col: str = "styleId",
    flag_col: str = "AI Generated Image Flag"
) -> ExtractMissingProcessorResult:
    """
    Process an extract missing data operation for Streamlit.

    This function handles the complete workflow:
    1. Receives raw file data from uploaded files
    2. Creates temporary files for processing
    3. Runs the extract missing formatter
    4. Returns the result as a BytesIO object for download

    Args:
        input_file_data: Raw bytes of the input Excel file (with Values and Types sheets)
        input_filename: Original filename of the input file
        output_file_data: Raw bytes of the output Excel file (with AI Generated Image Flag column)
        output_filename: Original filename of the output file
        result_filename: Name for the result file (default: "styles_with_missing_ai_flag.xlsx")
        values_sheet: Name of the Values sheet (default: "Values")
        types_sheet: Name of the Types sheet (default: "Types")
        style_id_col: Name of the style ID column (default: "styleId")
        flag_col: Name of the AI flag column (default: "AI Generated Image Flag")

    Returns:
        ExtractMissingProcessorResult with operation details and downloadable data
    """
    logger.info("Starting extract missing data processing", extra_data={
        "input_filename": input_filename,
        "output_filename": output_filename,
        "result_filename": result_filename
    })

    # Validate file sizes
    error = validate_file_size(input_file_data, input_filename)
    if error:
        logger.warning("Input file size validation failed", extra_data={"filename": input_filename, "error": error})
        return ExtractMissingProcessorResult(success=False, error_message=error)

    error = validate_file_size(output_file_data, output_filename)
    if error:
        logger.warning("Output file size validation failed", extra_data={"filename": output_filename, "error": error})
        return ExtractMissingProcessorResult(success=False, error_message=error)

    # Validate filename extensions
    if not input_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid input file extension", extra_data={"filename": input_filename})
        return ExtractMissingProcessorResult(
            success=False,
            error_message=f"Input file must be an Excel file (.xlsx or .xls)"
        )

    if not output_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid output file extension", extra_data={"filename": output_filename})
        return ExtractMissingProcessorResult(
            success=False,
            error_message=f"Output file must be an Excel file (.xlsx or .xls)"
        )

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Create temp file paths
            input_path = tmpdir / input_filename
            output_path = tmpdir / output_filename
            result_path = tmpdir / result_filename

            # Write uploaded data to temp files
            with open(input_path, "wb") as f:
                f.write(input_file_data)
            with open(output_path, "wb") as f:
                f.write(output_file_data)

            # Process the files
            result = extract_rows_with_missing_ai_flag(
                input_file=input_path,
                output_file=output_path,
                result_file=result_path,
                values_sheet=values_sheet,
                types_sheet=types_sheet,
                style_id_col=style_id_col,
                flag_col=flag_col
            )

            if result.success:
                logger.info("Extract missing data processing completed successfully", extra_data={
                    "rows_extracted": result.rows_extracted,
                    "types_rows": result.types_rows,
                    "missing_count": result.missing_count
                })
                # Read the output into BytesIO for download
                with open(result_path, "rb") as f:
                    output_data = BytesIO(f.read())

                return ExtractMissingProcessorResult(
                    success=True,
                    data=output_data,
                    filename=result_filename,
                    rows_extracted=result.rows_extracted,
                    types_rows=result.types_rows,
                    missing_count=result.missing_count
                )
            else:
                logger.warning("Extract missing data processing failed", extra_data={"error": result.error_message})
                return ExtractMissingProcessorResult(
                    success=False,
                    error_message=result.error_message
                )

    except Exception as e:
        logger.error("Unexpected error during extract missing data processing", extra_data={"error": str(e)})
        return ExtractMissingProcessorResult(
            success=False,
            error_message=f"Processing error: {e}"
        )
