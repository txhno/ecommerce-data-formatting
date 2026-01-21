"""
Command-line interface for Rubick CAST Formatting.
"""

from src.utils.logging import get_logger, setup_logging

logger = get_logger(__name__)


def main():
    """Main CLI entry point."""
    # Initialize logging for CLI (human-readable format)
    setup_logging(level="INFO", json_format=False)

    logger.info("Rubick CAST Formatting CLI initialized")
    print("Rubick CAST Formatting CLI")
    print("Usage: rubick-format <command> <args>")
    print("")
    print("Commands:")
    print("  export <input> <template> <output>  - Format input to match template")
    print("  import <sku> <style> <output>       - Merge SKU and Style files")
    print("")
    print("Use 'python -m streamlit run app.py' for the web interface.")


if __name__ == "__main__":
    main()
