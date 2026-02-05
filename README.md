# Golf Swing Statistics Analyzer

A Python tool for analyzing golf swing statistics from CSV data. This script processes swing data, allows selective exclusion of outlier shots, and generates a formatted Excel report with comprehensive statistics including averages and standard deviations.

## Features

- **CSV to Excel Conversion**: Transforms raw CSV swing data into a professionally formatted Excel workbook
- **Selective Data Exclusion**: Mark specific shots to exclude from statistical calculations while keeping them visible
- **Automatic Statistics**: Calculates averages and standard deviations for all numeric metrics
- **Dynamic Formulas**: Uses Excel formulas that automatically update when you toggle row inclusion
- **Professional Formatting**: Color-coded headers and statistics rows for easy readability
- **Auto-sizing Columns**: Automatically adjusts column widths for optimal viewing

## Requirements

```bash
pip install pandas numpy openpyxl
```

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/golf-stats-analyzer.git
cd golf-stats-analyzer
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Run the script interactively:

```bash
python golf_stats_analyzer.py
```

The script will prompt you for:
1. Input CSV file name
2. Output Excel file name
3. Row numbers to exclude (optional)

### Example

```
Enter input CSV file name: swingcaddie_6I.csv
Enter output XLSX file name: analysis.xlsx
Enter row numbers to exclude, separated by commas: 3,7,12
```

### Programmatic Usage

You can also import and use the function directly:

```python
from golf_stats_analyzer import process_golf_stats

process_golf_stats(
    input_file='swingcaddie_6I.csv',
    output_file='golf_stats_analysis.xlsx',
    excluded_rows=[3, 7, 12]
)
```

## Input File Format

Your CSV file should have the following structure:

- `No.`: Shot number (required)
- `Date`: Date of the shot (required)
- `EQ`: Equipment used (required)
- Additional numeric columns: Any swing metrics (Ball Speed, Club Speed, Carry, etc.)

Example CSV:
```csv
No.,Date,EQ,Ball Speed,Club Speed,Carry,Total,Smash Factor
1,2024-01-20,6I,115.2,89.3,165,175,1.29
2,2024-01-20,6I,117.8,90.1,168,178,1.31
...
```

## Output Format

The generated Excel file includes:

### Data Columns
- **No.**: Original shot number
- **Date**: Shot date
- **EQ**: Equipment used
- **Include**: Yes/No indicator (can be manually changed in Excel)
- **Metric columns**: All numeric data from the CSV

### Statistics Rows
- **AVG row** (yellow): Averages of all included shots
- **STDEV row** (red): Standard deviations of all included shots

### Interactive Features
- Change any "Include" value from "Yes" to "No" in Excel to exclude that shot
- Statistics automatically recalculate based on included shots
- All formulas use AVERAGEIF and conditional STDEV calculations

## How It Works

1. **Data Loading**: Reads CSV file and removes any existing AVG rows
2. **Exclusion Tracking**: Adds an "Include" column to track which rows to include in statistics
3. **Excel Generation**: Creates formatted workbook with:
   - Blue headers
   - Data rows with include/exclude flags
   - Yellow AVG row with AVERAGEIF formulas
   - Red STDEV row with conditional standard deviation formulas
4. **Formatting**: Auto-sizes columns and applies color coding

## Excel Formula Details

### Average Calculation
```excel
=AVERAGEIF($D$2:$D$N,"Yes",E2:E_N)
```
Calculates average only for rows where Include = "Yes"

### Standard Deviation Calculation
```excel
=SQRT(SUMPRODUCT(((E2:E_N-E_AVG)^2),--($D$2:$D$N="Yes"))/(SUMPRODUCT(--($D$2:$D$N="Yes"))-1))
```
Calculates population standard deviation for included rows only

## Use Cases

- **Golf Practice Analysis**: Track and analyze practice session statistics
- **Club Fitting**: Compare performance across different clubs
- **Swing Improvement**: Monitor progress over time with outlier removal
- **Launch Monitor Data**: Process data from devices like SwingCaddie, TrackMan, etc.

## Customization

You can modify the script to:
- Change color schemes by adjusting `PatternFill` colors
- Add additional statistics (median, min, max, etc.)
- Modify column width calculations
- Add conditional formatting rules

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

Your Name - [Your GitHub Profile](https://github.com/yourusername)

## Acknowledgments

- Built with pandas, numpy, and openpyxl
- Designed for golf swing analysis with launch monitors