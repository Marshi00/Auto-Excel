import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('Auto Template.xlsx', sheet_name=0)
""" 
# Specify start and end rows
start_row = 1
end_row = 52
"""
# Extract device type and name based on your criteria
df['location'] = df['Tag Name'].str.split('-').str[1]
df['Device_Type'] = df['Tag Name'].str.split('-').str[2]
df['Device_Name'] = df['Tag Name'].str.split('-').str[3]
"""
# Filter the data based on start and end rows
filtered_df = df.iloc[start_row - 1:end_row]
"""

# Drop empty cells (NaN values) in both 'Device_Type' and 'Device_Name' columns
filtered_df = df.dropna(subset=['Device_Type', 'Device_Name', 'location'])

# Remove duplicates based on both 'Device_Type' and 'Device_Name' columns
filtered_df = filtered_df.drop_duplicates(subset=['Device_Type', 'Device_Name', 'location'])

# Create a new DataFrame with extracted data
result_df = filtered_df[['Device_Type', 'Device_Name', 'location']]

# Save the result to a new Excel file or sheet
result_df.to_excel('rdy_device.xlsx', index=False)


# Read the 'pumps_template' Excel file
template_df = pd.read_excel('p.xlsx', sheet_name='Sheet1', header=None)
template_df.columns = ["BLOCK TYPE", "TAG", "DESCRIPTION"]
# Define the placeholders to be replaced
placeholder_device_type = "PLACEHOLDERDEVICETYPE"
placeholder_device_name = "PLACEHOLDERDEVICENAME"
placeholder_device_loc = "PLACEHOLDERLOCATION"


# Create an empty DataFrame to store the updated rows
updated_rows = []


# Iterate over rows in the 'pumps_template' DataFrame
for index, row in template_df.iterrows():
    # Replace placeholders with device type and device name
    updated_row = row.apply(lambda cell: cell.replace(placeholder_device_type, result_df['Device_Type'].values[0]))
    updated_row = updated_row.apply(lambda cell: cell.replace(placeholder_device_loc, result_df['location'].values[0]))
    updated_row = updated_row.apply(
        lambda cell: cell.replace(placeholder_device_name, result_df['Device_Name'].values[0]))

    updated_rows.append(updated_row)
""" 
mix same sheet with others later 
# Concatenate the updated rows into a new DataFrame
result_df = pd.concat(updated_rows, axis=1).transpose()
"""
final_df = pd.concat(updated_rows, axis=1).transpose()
# Save the result to a new Excel file or sheet
final_df.to_excel('output.xlsx', index=False)