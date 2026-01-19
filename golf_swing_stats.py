import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

def process_golf_stats(input_file, output_file, excluded_rows=None):
    """
    Process golf swing statistics CSV and generate XLSX with statistics.
    
    Parameters:
    - input_file: Path to input CSV file
    - output_file: Path to output XLSX file
    - excluded_rows: List of row numbers to exclude (1-indexed, matching the 'No.' column)
    """
    
    # Read the CSV file
    df = pd.read_csv(input_file)
    
    # Remove the existing AVG row if present
    df = df[df['No.'] != 'AVG'].copy()
    
    # Add an 'Include' column for tracking which rows to include
    df['Include'] = True
    
    # Exclude specified rows
    if excluded_rows:
        for row_num in excluded_rows:
            df.loc[df['No.'] == row_num, 'Include'] = False
    
    # Identify numeric columns (excluding No., Date, EQ, Include)
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if 'No.' in numeric_cols:
        numeric_cols.remove('No.')
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Golf Stats"
    
    # Write headers
    headers = ['No.', 'Date', 'EQ', 'Include'] + numeric_cols
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    # Write data rows
    row_idx = 2
    for _, row in df.iterrows():
        ws.cell(row=row_idx, column=1, value=row['No.'])
        ws.cell(row=row_idx, column=2, value=row['Date'])
        ws.cell(row=row_idx, column=3, value=row['EQ'])
        ws.cell(row=row_idx, column=4, value='Yes' if row['Include'] else 'No')
        
        for col_idx, col_name in enumerate(numeric_cols, 5):
            ws.cell(row=row_idx, column=col_idx, value=row[col_name])
        
        row_idx += 1
    
    # Add AVG row
    avg_row = row_idx
    ws.cell(row=avg_row, column=1, value='AVG')
    ws.cell(row=avg_row, column=2, value='')
    ws.cell(row=avg_row, column=3, value='')
    ws.cell(row=avg_row, column=4, value='')
    
    # Add formulas for averages
    for col_idx, col_name in enumerate(numeric_cols, 5):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        # AVERAGEIF formula to only include rows where Include='Yes'
        formula = f'=AVERAGEIF($D$2:$D${avg_row-1},"Yes",{col_letter}2:{col_letter}{avg_row-1})'
        ws.cell(row=avg_row, column=col_idx, value=formula)
    
    # Style AVG row
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=avg_row, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    
    # Add STDEV row
    stdev_row = avg_row + 1
    ws.cell(row=stdev_row, column=1, value='STDEV')
    ws.cell(row=stdev_row, column=2, value='')
    ws.cell(row=stdev_row, column=3, value='')
    ws.cell(row=stdev_row, column=4, value='')
    
    # Add formulas for standard deviation
    for col_idx, col_name in enumerate(numeric_cols, 5):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        # Create array formula for conditional STDEV
        # Using STDEV.S with IF array formula
        formula = f'=STDEV.S(IF($D$2:$D${avg_row-1}="Yes",{col_letter}2:{col_letter}{avg_row-1}))'
        ws.cell(row=stdev_row, column=col_idx, value=formula)
    
    # Style STDEV row
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=stdev_row, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    # Adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 20)
        ws.column_dimensions[column].width = adjusted_width
    
    # Save workbook
    wb.save(output_file)
    print(f"✓ Generated {output_file}")
    print(f"✓ Processed {len(df)} rows")
    print(f"✓ Excluded {sum(~df['Include'])} rows")
    print(f"✓ Included {sum(df['Include'])} rows in statistics")


# Example usage
if __name__ == "__main__":
    # Exclude row 8 (the outlier with very low values)
    excluded_rows = [8]

    process_golf_stats(
        # Prompt for input file
        input_file = input("Enter the path to the input CSV file: ").strip(),

        # Prompt for output file
        output_file = input("Enter the name for the output Excel file: ").strip(),

        excluded_rows=excluded_rows
    )
    
    print("\nTo exclude different rows, modify the 'excluded_rows' list.")
    print("For example: excluded_rows = [3, 8] to exclude rows 3 and 8")