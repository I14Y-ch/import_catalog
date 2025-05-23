# I14Y Dataset Import Tool

## Metadata Import Made Easy

Streamline your dataset management workflow! This tool enables efficient bulk importation of dataset metadata into Switzerland's I14Y platform. Using the Excel template and Python scripts, you can quickly transfer your dataset descriptions to I14Y without manual entry through the web interface. Perfect for organizations managing multiple datasets that need to be published on the Swiss interoperability platform.

## Features

- Import datasets from Excel files to I14Y API supporting the following DCAT fields:
  - Title and Description (monolingual)
  - Identifiers
  - Access Rights
  - Publication Dates
  - Keywords (up to 3, monolingual)
  - Contact Points
  - Themes
  - Spatial Coverage
  - Temporal Coverage
  - Distributions (up to 3)
- Generation of Excel templates for dataset input with data validation

## Prerequisites

- Python 3.8+
- pip package manager

## Usage

### Import Datasets

1. Grab the [Excel example](data/datasets.xlsx) (located in the data directory). Fill in the information about your datasets. Please don't modify the structure of the table. (If you do so, you might have to adapt the Python import script later.) 
2. Log in on the interoperability platform. Copy the token clicking on the profile symbol. Fill in the token in the file config.py. Also provide the identifier of your organsation. 
3. Run the import script:

```bash
python src/import_datasets.py
```

The script will process each row and display real-time progress and error messages in the terminal.

### Generate Template
### Generate Template

Normally, it is not necessary to create a new Excel template. The script `create_template.py` is primarily useful if you need to make major adjustments to the template and import script, or if controlled vocabularies need to be updated.

To generate a new template:

```bash
python src/create_template.py
```

This will generate `data/template_datasets.xlsx` with data validation and reference sheets.

## Installation

1. Clone this repository:
```bash
git clone [repository-url]
cd bfs/2025_02_batchimport
```

2. (Optional but recommended) Create and activate a virtual environment:
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Configure the application:
   - Edit `src/config.py` with your I14Y API token and organization ID


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