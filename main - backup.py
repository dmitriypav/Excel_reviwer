import pandas as pd

# Load Excel file
file_path = 'Data/file_name.xlsx'
df = pd.read_excel(file_path)
print(df)

# Add new column
df['Result'] = ''

# Extract value from cell B2
value_b2 = df.iloc[0, 1]
# Verification Address
if pd.notna(value_b2) and value_b2 != '':
    df.at[0, 'Result'] = 'OK'
else:
    df.at[0, 'Result'] = f'FAIL {value_b2}'

# Value B3
value_b3 = df.iloc[1, 1]
# Value B4
value_b4 = df.iloc[2, 1]

if value_b3 == 'FortiNet' and value_b4 == 'FortiGate 70D':
    df.at[1, 'Result'] = (f'OK'
                          f'\nFirewall model {value_b4}')
else:
    df.at[1, 'Result'] = (f'FALSE'
                          f'\nFirewall model {value_b4}')
# print(df)

# Save to file
output_path = 'Data/file_name-verified.xlsx'
df.to_excel(output_path, index=False)

print('File saved:', output_path)