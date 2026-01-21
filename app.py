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

    col1, col2 = st.columns(2, gap="large")

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
        if st.button("Start Export", use_container_width=True, type="primary"):
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
        if st.button("Start Import", use_container_width=True, type="primary"):
            st.session_state.current_mode = "import"
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


if __name__ == "__main__":
    main()
