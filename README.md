# License Summary Tool

A Python script for analyzing and summarizing user license requirements in Dynamics 365 F&O.

## Features

- Analyzes user security roles and their license requirements
- Reads roles and their license requirements from a configuration file
- Generates summaries of unique role combinations
- Calculates required licenses for each combination
- Creates a formatted Excel report with the results

## Prerequisites

- Python 3.x
- Required Python packages (see `requirements.txt`)

## Installation

1. Clone this repository or download the script
2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

Run the script with both the license report and roles file:
```bash
python license_summary.py "License Report.xlsx" "Roles.xlsx"
```

### Input Files

1. Dynamics License Report:
   - An Excel file containing user security roles
   - Expected format: Standard Dynamics 365 F&O license report
   - Security roles are expected in column F

2. Roles File (required):
   - An Excel file containing role definitions and license requirements
   - Format:
     - Column A: Role names
     - Column B: Finance license (1 if required)
     - Column C: SCM license (1 if required)
     - Column D: Commerce license (1 if required)
     - Column E: Project license (1 if required)
     - Column F: HR license (1 if required)

### Output

The script generates a new Excel file with the following information:
- Count of users for each role combination
- Individual license requirements
- Combined license requirements
- Summary totals

The output file will be named: `[original_filename]_summary.xlsx` 