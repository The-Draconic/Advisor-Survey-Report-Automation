import pandas as pd

# Read data from the Excel sheet
excel_file = 'sample_data.xlsx'  # Replace with your Excel file name
df = pd.read_excel(excel_file)

# Specify the condition for splitting (e.g., based on a column value)
condition_column = 'Category'  # Replace with the column name for the condition
unique_values = df[condition_column].unique()

# Create separate Excel sheets for each unique value in the specified column
for value in unique_values:
    filtered_df = df[df[condition_column] == value]
    
    # Write the filtered data to a new Excel sheet
    output_excel_file = f'{value}_sheet.xlsx'  # Output file name based on the condition
    filtered_df.to_excel(output_excel_file, index=False)

print("Data has been successfully split into separate Excel sheets.")
