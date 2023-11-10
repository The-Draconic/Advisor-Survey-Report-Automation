import pandas as pd

# Sample Data
data = {
    'Name': ['John', 'Alice', 'Bob', 'Eve', 'Charlie'],
    'Category': ['A', 'B', 'A', 'C', 'B'],
    'Value': [10, 20, 15, 25, 30]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Save to Excel
df.to_excel('sample_data.xlsx', index=False)
