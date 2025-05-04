# Demand Normalization Script - Usage Guide

## Overview
This script processes an Excel file containing demand data according to the requirements specified in the README.md file. It performs two main operations:
1. Combines cells with values less than 30 with the next cell to the right
2. Rounds values greater than 0 to the nearest multiple of 10

## Requirements
- Python 3.6 or higher
- Required Python packages:
  - pandas
  - numpy

You can install the required packages using pip:
```
pip install pandas numpy
```

## How to Use

### Basic Usage
Run the script with the default input file (SV.xlsx):
```
python demand_normalization.py
```

### Specify a Different Input File
You can specify a different input file as a command-line argument:
```
python demand_normalization.py path/to/your/file.xlsx
```

## Output
The script will create a new Excel file with "_output" appended to the original filename. For example, if the input file is "SV.xlsx", the output file will be "SV_output.xlsx".

The output file will contain:
- The processed "Demand" sheet with the combined and rounded values
- All other sheets from the original file, unchanged

## Notes
- The script looks for a sheet named "Demand" (case-insensitive) in the Excel file
- The script processes data starting from row 2 (assuming row 1 contains headers)
- The script processes data starting from column C onwards (assuming columns A and B contain item code and description)
