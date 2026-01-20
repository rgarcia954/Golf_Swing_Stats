import pandas as pd
import xlsxwriter
import xlsxwriter.utility as xl_util
import os

def process_swing_caddie_data(input_csv, output_xlsx):
    """
    Processes the Swing Caddie CSV and generates an Excel file with:
    1. A new 'I' toggle column.
    2. Dynamic IF formulas for each data cell.
    3. Live AVERAGEIF formulas for non-zero averages.
    """
    # 1. Load the data
    df = pd.read_csv(input_csv)

    # 2. Clean data: Remove existing average row if present
    # Usually identified by 'AVG' in the first column
    first_col_name = df.columns[0]
    df = df[df[first_col_name].astype(str).str.strip().str.upper() != 'AVG'].reset_index(drop=True)

    # 3. Insert new column 'I' after the first column (Excel Column B)
    df.insert(1, 'I', 0)

    # 4. Create the Excel file using xlsxwriter
    writer = pd.ExcelWriter(output_xlsx, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Practice')
    
    workbook = writer.book
    worksheet = writer.sheets['Practice']
    
    # Define formatting
    avg_format = workbook.add_format({'bold': True, 'top': 2, 'num_format': '#,##0.00'})

    num_rows = len(df)
    num_cols = len(df.columns)

    # 5. Overwrite the data rows with IF formulas
    for row_idx in range(num_rows):
        excel_row_num = row_idx + 2  # Excel is 1-indexed, +1 for header
        
        for col_idx in range(2, num_cols):
            original_val = df.iloc[row_idx, col_idx]
            i_cell_ref = f"$B{excel_row_num}"
            
            # Format formula based on data type
            if isinstance(original_val, str):
                formula = f'=IF({i_cell_ref}=1, 0, "{original_val}")'
            elif pd.isna(original_val):
                formula = f'=IF({i_cell_ref}=1, 0, 0)'
            else:
                formula = f'=IF({i_cell_ref}=1, 0, {original_val})'
            
            worksheet.write_formula(excel_row_num - 1, col_idx, formula)

    # 6. Add the Average row at the bottom
    avg_row_idx = num_rows + 1
    worksheet.write(avg_row_idx, 0, 'AVG', avg_format)
    
    for col_idx in range(1, num_cols):
        col_letter = xl_util.xl_col_to_name(col_idx)
        data_range = f"{col_letter}2:{col_letter}{num_rows + 1}"
        
        # Formula: Average only if value is not zero
        formula = f'=AVERAGEIF({data_range}, "<>0")'
        worksheet.write_formula(avg_row_idx, col_idx, formula, avg_format)

    writer.close()

def main():
    print("--- Swing Caddie Data Processor ---")
    
    # Prompt for input file
    input_file = input("Enter the path to the input CSV file (e.g., practice.csv): ").strip()
    
    # Check if file exists
    if not os.path.exists(input_file):
        print(f"Error: The file '{input_file}' was not found.")
        return

    # Prompt for output file
    output_file = input("Enter the name for the output Excel file (e.g., results.xlsx): ").strip()
    
    # Ensure the output has the correct extension
    if not output_file.lower().endswith('.xlsx'):
        output_file += '.xlsx'

    print(f"Processing '{input_file}'...")
    
    try:
        process_swing_caddie_data(input_file, output_file)
        print(f"Successfully generated: {output_file}")
    except Exception as e:
        print(f"An error occurred during processing: {e}")

if __name__ == "__main__":
    main()