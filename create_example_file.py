import pandas as pd

# Generate a sample dataframe with x rows and 5 columns
n_rows = 12500
data = {
    'A': range(1, n_rows+1),
    'B': ["Value_" + str(i) for i in range(1, n_rows+1)],
    'C': [i * 1.5 for i in range(1, n_rows+1)],
    'D': ["Text_" + str(i) for i in range(1, n_rows+1)],
    'E': [i * 2.5 for i in range(1, n_rows+1)]
}

df = pd.DataFrame(data)

# Save dataframe to an Excel file
df.to_excel("sample_12500_rows.xlsx", index=False)