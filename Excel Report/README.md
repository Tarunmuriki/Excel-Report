# Excel Report Generator

A Python application that generates Excel reports from CSV data with pivot tables, charts, and summary statistics.

## Features

- CSV file import
- Pivot table generation
- Chart creation using matplotlib
- Styled Excel export using openpyxl
- User-friendly GUI interface
- Summary statistics

## Requirements

- Python 3.8+
- pandas
- openpyxl
- matplotlib
- tkinter (usually comes with Python)

## Installation

1. Clone this repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python src/main.py
   ```
2. Use the GUI to:
   - Select input CSV file
   - Choose report options
   - Generate and save Excel report

## Project Structure

```
.
├── src/
│   ├── main.py            # Main application with GUI
│   ├── report_generator.py # Report generation logic
│   └── utils.py           # Utility functions
├── data/
│   └── sample_data.csv    # Sample data
└── requirements.txt       # Project dependencies
```