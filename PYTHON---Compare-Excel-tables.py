import pandas as pd

def compare_excel_tables(file1, sheet1, file2, sheet2, output_file):
    """
    Compares two Excel sheets and saves the differences to a new Excel file.

    :param file1: Path to the first Excel file.
    :param sheet1: Sheet name or index of the first Excel file.
    :param file2: Path to the second Excel file.
    :param sheet2: Sheet name or index of the second Excel file.
    :param output_file: Path to save the differences.
    """
    # Load the Excel files
    df1 = pd.read_excel(file1, sheet_name=sheet1)
    df2 = pd.read_excel(file2, sheet_name=sheet2)

    # Ensure both DataFrames have the same columns
    if set(df1.columns) != set(df2.columns):
        raise ValueError("The two tables have different column structures.")

    # Reorder columns to match
    df2 = df2[df1.columns]

    # Compare the DataFrames
    comparison = df1.compare(df2, align_axis=0)

    # Save the differences to an Excel file
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        comparison.to_excel(writer, sheet_name='Differences')

    print(f"Comparison completed. Differences saved to '{output_file}'.")

# Example usage
if __name__ == "__main__":
    file1 = "table1.xlsx"  # Path to the first Excel file
    sheet1 = 0             # Sheet index or name (e.g., 0 or 'Sheet1')
    file2 = "table2.xlsx"  # Path to the second Excel file
    sheet2 = 0             # Sheet index or name (e.g., 0 or 'Sheet1')
    output_file = "differences.xlsx"  # Path to save the differences

    compare_excel_tables(file1, sheet1, file2, sheet2, output_file)
