# I14Y Dataset Import Tool

A Python-based tool for creating and importing datasets into the I14Y platform of the Swiss Federal Statistical Office (BFS).

## Features

- Generate Excel templates for dataset input with data validation
- Import datasets from Excel files to I14Y API
- Support for required DCAT fields:
  - Title and Description (DE)
  - Identifiers
  - Access Rights
  - Publication Dates
- Support for optional fields:
  - Keywords
  - Contact Points
  - Themes
  - Spatial Coverage
  - Temporal Coverage
  - Distributions (up to 3)

## Prerequisites

- Python 3.8+
- pip package manager

## Installation

1. Clone this repository:
```bash
git clone [repository-url]
cd bfs/2025_02_batchimport
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Configure the application:
   - Edit `src/config.py` with your I14Y API token and organization ID

## Usage

### Generate Template

Create a new Excel template for data input:

```bash
python src/create_template.py
```

This will generate `data/template_datasets.xlsx` with data validation and reference sheets.

### Import Datasets

1. Fill in the Excel template with your dataset information
2. Run the import script:

```bash
python src/import_datasets.py
```

The script will process each row and display real-time progress and error messages in the terminal.

## File Structure

```
2025_02_batchimport/
├── data/
│   └── datasets.xlsx
├── src/
│   ├── config.py
│   ├── create_template.py
│   └── import_datasets.py
├── requirements.txt
└── README.md
```

## Contributing

Please ensure any pull requests or contributions adhere to the following guidelines:
- Keep the code simple and well-documented
- Follow PEP 8 style guidelines
- Include appropriate error handling
- Test thoroughly before submitting
