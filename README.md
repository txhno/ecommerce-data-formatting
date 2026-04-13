# Excel Data Formatting

A tool for formatting Excel files for template-based export and import workflows.

## Features

- **Export Format**: Reindex input data to match a template's column structure
- **Import Merge**: Combine size chart (SKU) and product details (Style) files into a unified workbook format

## Installation

```bash
# Create virtual environment
uv venv venv -python 3.12.0

# Activate virtual environment
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install package
uv pip install -e .
```

## Usage

### Streamlit Web Interface

```bash
python -m streamlit run app.py
```

### Command Line

```bash
python -m src.cli
```

## Configuration

Copy `.env.example` to `.env` and customize settings:

```bash
cp .env.example .env
```

## Project Structure

```
ecommerce-data-formatting/
├── app.py                 # Streamlit web application
├── pyproject.toml         # Project configuration
├── src/
│   ├── __init__.py
│   ├── config.py          # Configuration management
│   ├── core/
│   │   ├── __init__.py
│   │   ├── export_formatter.py   # Export formatting logic
│   │   └── import_formatter.py   # Import merging logic
│   ├── ui/
│   │   └── __init__.py
│   └── utils/
│       └── __init__.py
├── config/                # Configuration files
├── tests/                 # Test files
└── templates/             # Template files
```
