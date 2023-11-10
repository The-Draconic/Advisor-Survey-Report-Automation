import pandas as pd
import os

def test_read_data():
    excel_file = 'sample_data.xlsx'
    df = pd.read_excel(excel_file)
    assert not df.empty, "Failed to read data from the Excel sheet."

def test_condition_splitting():
    excel_file = 'sample_data.xlsx'
    df = pd.read_excel(excel_file)
    condition_column = 'Category'
    unique_values = df[condition_column].unique()

    for value in unique_values:
        output_excel_file = f'{value}_sheet.xlsx'
        filtered_df = df[df[condition_column] == value]
        filtered_df.to_excel(output_excel_file, index=False)
        assert not filtered_df.empty, f"Failed to split data for {value} in the condition column."

def test_output_files_exist():
    excel_file = 'sample_data.xlsx'
    df = pd.read_excel(excel_file)
    condition_column = 'Category'
    unique_values = df[condition_column].unique()

    for value in unique_values:
        output_excel_file = f'{value}_sheet.xlsx'
        assert os.path.exists(output_excel_file), f"Output file {output_excel_file} does not exist."

def test_output_files_not_empty():
    excel_file = 'sample_data.xlsx'
    df = pd.read_excel(excel_file)
    condition_column = 'Category'
    unique_values = df[condition_column].unique()

    for value in unique_values:
        output_excel_file = f'{value}_sheet.xlsx'
        output_df = pd.read_excel(output_excel_file)
        assert not output_df.empty, f"Output file {output_excel_file} is empty."

# Run the tests
test_read_data()
test_condition_splitting()
test_output_files_exist()
test_output_files_not_empty()
print("All tests passed successfully.")
