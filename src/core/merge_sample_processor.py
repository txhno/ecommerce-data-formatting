"""
Merge sample output processor module for Streamlit integration.

Provides a clean interface for the web UI to process merge sample output operations
without dealing with file handling details.
"""

import tempfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Optional

from src.utils.logging import get_logger
from .merge_sample_formatter import merge_sample_output, MergeSampleResult

logger = get_logger(__name__)


@dataclass
class MergeSampleProcessorResult:
    """Result of a merge sample output processing operation for Streamlit."""
    success: bool
    data: Optional[BytesIO] = None
    filename: Optional[str] = None
    rows_updated: int = 0
    total_rows: int = 0
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


def process_merge_sample(
    output_file_data: bytes,
    output_filename: str,
    sample_file_data: bytes,
    sample_filename: str,
    result_filename: str = "output_rewritten.xlsx",
    style_id_col: str = "styleId"
) -> MergeSampleProcessorResult:
    """
    Process a merge sample output operation for Streamlit.

    This function handles the complete workflow:
    1. Receives raw file data from uploaded files
    2. Creates temporary files for processing
    3. Runs the merge sample formatter
    4. Returns the result as a BytesIO object for download

    Args:
        output_file_data: Raw bytes of the main output Excel file
        output_filename: Original filename of the output file
        sample_file_data: Raw bytes of the sample output Excel file
        sample_filename: Original filename of the sample file
        result_filename: Name for the result file (default: "output_rewritten.xlsx")
        style_id_col: Name of the style ID column (default: "styleId")

    Returns:
        MergeSampleProcessorResult with operation details and downloadable data
    """
    logger.info("Starting merge sample output processing", extra_data={
        "output_filename": output_filename,
        "sample_filename": sample_filename,
        "result_filename": result_filename
    })

    # Validate file sizes
    error = validate_file_size(output_file_data, output_filename)
    if error:
        logger.warning("Output file size validation failed", extra_data={"filename": output_filename, "error": error})
        return MergeSampleProcessorResult(success=False, error_message=error)

    error = validate_file_size(sample_file_data, sample_filename)
    if error:
        logger.warning("Sample file size validation failed", extra_data={"filename": sample_filename, "error": error})
        return MergeSampleProcessorResult(success=False, error_message=error)

    # Validate filename extensions
    if not output_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid output file extension", extra_data={"filename": output_filename})
        return MergeSampleProcessorResult(
            success=False,
            error_message=f"Output file must be an Excel file (.xlsx or .xls)"
        )

    if not sample_filename.lower().endswith(('.xlsx', '.xls')):
        logger.warning("Invalid sample file extension", extra_data={"filename": sample_filename})
        return MergeSampleProcessorResult(
            success=False,
            error_message=f"Sample file must be an Excel file (.xlsx or .xls)"
        )

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Create temp file paths
            output_path = tmpdir / output_filename
            sample_path = tmpdir / sample_filename
            result_path = tmpdir / result_filename

            # Write uploaded data to temp files
            with open(output_path, "wb") as f:
                f.write(output_file_data)
            with open(sample_path, "wb") as f:
                f.write(sample_file_data)

            # Process the files
            result = merge_sample_output(
                output_file=output_path,
                sample_file=sample_path,
                result_file=result_path,
                style_id_col=style_id_col
            )

            if result.success:
                logger.info("Merge sample output processing completed successfully", extra_data={
                    "rows_updated": result.rows_updated,
                    "total_rows": result.total_rows
                })
                # Read the output into BytesIO for download
                with open(result_path, "rb") as f:
                    output_data = BytesIO(f.read())

                return MergeSampleProcessorResult(
                    success=True,
                    data=output_data,
                    filename=result_filename,
                    rows_updated=result.rows_updated,
                    total_rows=result.total_rows
                )
            else:
                logger.warning("Merge sample output processing failed", extra_data={"error": result.error_message})
                return MergeSampleProcessorResult(
                    success=False,
                    error_message=result.error_message
                )

    except Exception as e:
        logger.error("Unexpected error during merge sample output processing", extra_data={"error": str(e)})
        return MergeSampleProcessorResult(
            success=False,
            error_message=f"Processing error: {e}"
        )
