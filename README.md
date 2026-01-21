# Rubick CAST Formatting

A tool for formatting Excel files to CAST import format.

## Features

- **Export Format**: Reindex input data to match a template's column structure
- **Import Merge**: Combine size chart (SKU) and product details (Style) files into CAST format

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
# Export format
rubick-format export input.xlsx template.xlsx output.xlsx

# Import merge
rubick-format import sku.xlsx style.xlsx output.xlsx
```

## Configuration

Copy `.env.example` to `.env` and customize settings:

```bash
cp .env.example .env
```

## Project Structure

```
rubick-cast-formatting/
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
