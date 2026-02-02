"""
Rubick CAST Formatting - Streamlit Application

A web interface for formatting Excel files to CAST import format.
Supports both export formatting (reindex to template) and
import merging (combine size chart and product details).
"""

import streamlit as st
from pathlib import Path

from src.config import get_settings
from src.core.export_processor import process_export, ExportProcessorResult
from src.core.import_processor import process_import, ImportProcessorResult
from src.core.extract_missing_processor import process_extract_missing, ExtractMissingProcessorResult
from src.core.merge_sample_processor import process_merge_sample, MergeSampleProcessorResult
from src.utils.logging import get_logger, setup_logging

logger = get_logger(__name__)


def setup_page():
    """Configure Streamlit page settings."""
    settings = get_settings()
    st.set_page_config(
        page_title=settings.app.app_name,
        page_icon="📊",
        layout="centered",
        initial_sidebar_state="collapsed"
    )


def render_logo():
    """Render the application logo if available."""
    logo_path = Path(__file__).parent / "logo.svg"
    if logo_path.exists():
        st.image(str(logo_path), width=120)


def render_landing_page():
    """Render the landing page with mode selection cards."""
    settings = get_settings()

    # Center the content
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        render_logo()
        st.title(f"{settings.app.app_name}")
        st.caption(f"Version {settings.app.app_version}")
        st.markdown("---")

    st.markdown("### Choose a Processing Mode")

    # 2x2 grid layout
    col1, col2 = st.columns(2, gap="large")
    col3, col4 = st.columns(2, gap="large")

    with col1:
        st.markdown("""
        <div class="mode-card export-card">
            <h2>📤 Export Format</h2>
            <p>Reindex your input data to match a template's column structure.</p>
            <ul>
                <li>Upload your data file</li>
                <li>Select a template</li>
                <li>Download formatted output</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Start Export", use_container_width=True, type="primary", key="export"):
            st.session_state.current_mode = "export"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="mode-card import-card">
            <h2>📥 Import Merge</h2>
            <p>Merge size chart (SKU) and product details (Style) files into CAST format.</p>
            <ul>
                <li>Combine SKU & Style files</li>
                <li>Auto-generate Types sheet</li>
                <li>Export merged dataset</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Start Import", use_container_width=True, type="primary", key="import"):
            st.session_state.current_mode = "import"
            st.rerun()

    with col3:
        st.markdown("""
        <div class="mode-card extract-missing-card">
            <h2>🔍 Extract Missing Data</h2>
            <p>Extract rows where AI Generated Image Flag is missing or empty.</p>
            <ul>
                <li>Upload input and output files</li>
                <li>Find rows with missing flags</li>
                <li>Download filtered results</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Extract Missing Data", use_container_width=True, type="primary", key="extract_missing"):
            st.session_state.current_mode = "extract_missing"
            st.rerun()

    with col4:
        st.markdown("""
        <div class="mode-card merge-sample-card">
            <h2>🔄 Merge Sample Output</h2>
            <p>Merge sample output updates into main output file by styleId.</p>
            <ul>
                <li>Upload main and sample files</li>
                <li>Auto-match by styleId</li>
                <li>Download updated output</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.button("Merge Sample Output", use_container_width=True, type="primary", key="merge_sample"):
            st.session_state.current_mode = "merge_sample"
            st.rerun()

    # Add custom CSS for mode cards
    st.markdown("""
    <style>
    .mode-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        text-align: center;
        height: 100%;
        border: 2px solid transparent;
        transition: all 0.3s ease;
    }
    .mode-card:hover {
        border-color: #4e8cff;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .mode-card h2 {
        margin-bottom: 10px;
        color: #1f2937;
    }
    .mode-card p {
        color: #6b7280;
        font-size: 0.9rem;
        margin-bottom: 15px;
    }
    .mode-card ul {
        text-align: left;
        color: #4b5563;
        font-size: 0.85rem;
        padding-left: 20px;
    }
    .export-card {
        border-left: 4px solid #10b981;
    }
    .import-card {
        border-left: 4px solid #6366f1;
    }
    .extract-missing-card {
        border-left: 4px solid #f59e0b;
    }
    .merge-sample-card {
        border-left: 4px solid #f43f5e;
    }
    </style>
    """, unsafe_allow_html=True)


def render_export_page():
    """Render the export formatting page."""
    settings = get_settings()

    # Header with back button
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.current_mode = None
            st.rerun()
    with col2:
        st.title("📤 Export Format")
        st.markdown("Reindex input data to match a template's column structure.")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        input_file = st.file_uploader(
            "Input Excel File",
            type=["xlsx", "xls"],
            key="export_input"
        )

    with col2:
        template_file = st.file_uploader(
            "Template Excel File",
            type=["xlsx", "xls"],
            key="export_template"
        )

    preserve_unknown = st.checkbox(
        "Preserve unknown columns",
        value=False,
        help="Keep columns from input that are not in the template"
    )

    max_size = settings.app.max_file_size_mb * 1024 * 1024

    # Validate file sizes
    if input_file and input_file.size > max_size:
        st.error(f"Input file exceeds maximum size ({settings.app.max_file_size_mb}MB)")
        input_file = None

    if template_file and template_file.size > max_size:
        st.error(f"Template file exceeds maximum size ({settings.app.max_file_size_mb}MB)")
        template_file = None

    if input_file and template_file:
        output_filename = st.text_input(
            "Output Filename",
            value="Formatted_Output.xlsx",
            help="Must end with .xlsx or .xls"
        )

        def validate_filename(filename: str) -> bool:
            if not filename:
                return False
            return filename.lower().endswith(('.xlsx', '.xls'))

        if output_filename and not validate_filename(output_filename):
            st.error("Output filename must end with .xlsx or .xls")
            output_filename = "Formatted_Output.xlsx"

        if st.button("Format Excel File", type="primary"):
            result = process_export(
                input_file_data=input_file.getvalue(),
                input_filename=input_file.name,
                template_file_data=template_file.getvalue(),
                template_filename=template_file.name,
                output_filename=output_filename,
                preserve_unknown_columns=preserve_unknown
            )

            if result.success:
                st.success("Format completed successfully!")

                with st.expander("Processing Details", expanded=True):
                    st.json({
                        "Rows Processed": result.rows_processed,
                        "Input Columns": result.columns_in_input,
                        "Output Columns": result.columns_in_output,
                        "Columns Added": result.columns_added
                    })

                st.download_button(
                    label="Download Formatted File",
                    data=result.data,
                    file_name=result.filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(f"Error: {result.error_message}")


def render_import_page():
    """Render the import merging page."""
    # Header with back button
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.current_mode = None
            st.rerun()
    with col2:
        st.title("📥 Import Merge")
        st.markdown("Merge size chart (SKU) and product details (Style) files into CAST format.")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        size_chart_file = st.file_uploader(
            "Size Chart File (sku.xlsx)",
            type=["xlsx", "xls"],
            key="import_size"
        )

    with col2:
        product_details_file = st.file_uploader(
            "Product Details File (style.xlsx)",
            type=["xlsx", "xls"],
            key="import_product"
        )

    output_filename = st.text_input(
        "Output Filename",
        value="Batch_Merged_With_Types_Values.xlsx",
        help="Must end with .xlsx or .xls"
    )

    def validate_filename(filename: str) -> bool:
        if not filename:
            return False
        return filename.lower().endswith(('.xlsx', '.xls'))

    if output_filename and not validate_filename(output_filename):
        st.error("Output filename must end with .xlsx or .xls")
        output_filename = "Batch_Merged_With_Types_Values.xlsx"

    exclude_sheets = st.text_input(
        "Exclude Sheets (comma-separated)",
        value="masterdata",
        help="Enter sheet names to exclude, separated by commas"
    )

    exclude_list = []
    if exclude_sheets.strip():
        exclude_list = [s.strip() for s in exclude_sheets.split(",") if s.strip()]
        invalid = [s for s in exclude_list if not s or any(c in s for c in r'[]/?*:;{}')]
        if invalid:
            st.error(f"Invalid sheet names (cannot contain special characters): {invalid}")
            exclude_list = []
            st.stop()

    if size_chart_file and product_details_file:
        if st.button("Merge Files", type="primary"):
            result = process_import(
                size_chart_data=size_chart_file.getvalue(),
                size_chart_filename=size_chart_file.name,
                product_details_data=product_details_file.getvalue(),
                product_details_filename=product_details_file.name,
                output_filename=output_filename,
                exclude_sheets=exclude_list
            )

            if result.success:
                st.success("Merge completed successfully!")

                with st.expander("Processing Details", expanded=True):
                    st.json({
                        "Rows Processed": result.rows_processed,
                        "Columns": result.columns_count,
                        "Sheets Processed": result.sheets_processed
                    })

                st.download_button(
                    label="Download Merged File",
                    data=result.data,
                    file_name=result.filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(f"Error: {result.error_message}")


def render_extract_missing_page():
    """Render the extract missing data page."""
    # Header with back button
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.current_mode = None
            st.rerun()
    with col2:
        st.title("🔍 Extract Missing Data")
        st.markdown("Extract rows where AI Generated Image Flag is missing or empty.")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        input_file = st.file_uploader(
            "Input Excel File (with Values and Types sheets)",
            type=["xlsx", "xls"],
            key="extract_input"
        )

    with col2:
        output_file = st.file_uploader(
            "Output Excel File (with AI Generated Image Flag column)",
            type=["xlsx", "xls"],
            key="extract_output"
        )

    output_filename = st.text_input(
        "Output Filename",
        value="styles_with_missing_ai_flag.xlsx",
        help="Must end with .xlsx or .xls"
    )

    def validate_filename(filename: str) -> bool:
        if not filename:
            return False
        return filename.lower().endswith(('.xlsx', '.xls'))

    if output_filename and not validate_filename(output_filename):
        st.error("Output filename must end with .xlsx or .xls")
        output_filename = "styles_with_missing_ai_flag.xlsx"

    max_size = get_settings().app.max_file_size_mb * 1024 * 1024

    # Validate file sizes
    if input_file and input_file.size > max_size:
        st.error(f"Input file exceeds maximum size ({get_settings().app.max_file_size_mb}MB)")
        input_file = None

    if output_file and output_file.size > max_size:
        st.error(f"Output file exceeds maximum size ({get_settings().app.max_file_size_mb}MB)")
        output_file = None

    if input_file and output_file:
        if st.button("Extract Missing Data", type="primary"):
            result = process_extract_missing(
                input_file_data=input_file.getvalue(),
                input_filename=input_file.name,
                output_file_data=output_file.getvalue(),
                output_filename=output_file.name,
                result_filename=output_filename
            )

            if result.success:
                st.success("Extraction completed successfully!")

                with st.expander("Processing Details", expanded=True):
                    st.json({
                        "Rows Extracted": result.rows_extracted,
                        "Missing Flags Found": result.missing_count,
                        "Types Rows": result.types_rows
                    })

                st.download_button(
                    label="Download Extracted File",
                    data=result.data,
                    file_name=result.filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(f"Error: {result.error_message}")


def render_merge_sample_page():
    """Render the merge sample output page."""
    # Header with back button
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("← Back", use_container_width=True):
            st.session_state.current_mode = None
            st.rerun()
    with col2:
        st.title("🔄 Merge Sample Output")
        st.markdown("Merge sample output updates into main output file by styleId.")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        main_output_file = st.file_uploader(
            "Main Output Excel File",
            type=["xlsx", "xls"],
            key="merge_main"
        )

    with col2:
        sample_output_file = st.file_uploader(
            "Sample Output Excel File",
            type=["xlsx", "xls"],
            key="merge_sample"
        )

    output_filename = st.text_input(
        "Output Filename",
        value="output_rewritten.xlsx",
        help="Must end with .xlsx or .xls"
    )

    def validate_filename(filename: str) -> bool:
        if not filename:
            return False
        return filename.lower().endswith(('.xlsx', '.xls'))

    if output_filename and not validate_filename(output_filename):
        st.error("Output filename must end with .xlsx or .xls")
        output_filename = "output_rewritten.xlsx"

    max_size = get_settings().app.max_file_size_mb * 1024 * 1024

    # Validate file sizes
    if main_output_file and main_output_file.size > max_size:
        st.error(f"Main output file exceeds maximum size ({get_settings().app.max_file_size_mb}MB)")
        main_output_file = None

    if sample_output_file and sample_output_file.size > max_size:
        st.error(f"Sample output file exceeds maximum size ({get_settings().app.max_file_size_mb}MB)")
        sample_output_file = None

    if main_output_file and sample_output_file:
        if st.button("Merge Sample Data", type="primary"):
            result = process_merge_sample(
                output_file_data=main_output_file.getvalue(),
                output_filename=main_output_file.name,
                sample_file_data=sample_output_file.getvalue(),
                sample_filename=sample_output_file.name,
                result_filename=output_filename
            )

            if result.success:
                st.success("Merge completed successfully!")

                with st.expander("Processing Details", expanded=True):
                    st.json({
                        "Rows Updated": result.rows_updated,
                        "Total Rows": result.total_rows
                    })

                st.download_button(
                    label="Download Merged File",
                    data=result.data,
                    file_name=result.filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(f"Error: {result.error_message}")


def main():
    """Main application entry point."""
    settings = get_settings()
    logger.info("Starting Rubick CAST Formatting application", extra_data={
        "app_name": settings.app.app_name,
        "app_version": settings.app.app_version
    })

    setup_page()

    # Initialize session state for mode tracking
    if "current_mode" not in st.session_state:
        st.session_state.current_mode = None

    # Render appropriate page based on current mode
    current_mode = st.session_state.current_mode

    if current_mode is None:
        render_landing_page()
    elif current_mode == "export":
        render_export_page()
    elif current_mode == "import":
        render_import_page()
    elif current_mode == "extract_missing":
        render_extract_missing_page()
    elif current_mode == "merge_sample":
        render_merge_sample_page()


if __name__ == "__main__":
    main()
