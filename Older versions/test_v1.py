import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('Auto Template.xlsx', sheet_name=0)

# Extract device type and name based on your criteria
df['location'] = df['Tag Name'].str.split('-').str[1]
df['Device_Type'] = df['Tag Name'].str.split('-').str[2]
df['Device_Name'] = df['Tag Name'].str.split('-').str[3]

# Drop empty cells (NaN values) in both 'Device_Type' and 'Device_Name' columns
filtered_df = df.dropna(subset=['Device_Type', 'Device_Name', 'location'])

# Remove duplicates based on both 'Device_Type' and 'Device_Name' columns
filtered_df = filtered_df.drop_duplicates(subset=['Device_Type', 'Device_Name', 'location'])

# Create a new DataFrame with extracted data
result_df = filtered_df[['Device_Type', 'Device_Name', 'location']]

# Save the result to a new Excel file or sheet
result_df.to_excel('rdy_device2.xlsx', index=False)

# Define the placeholders to be replaced
placeholder_device_type = "PLACEHOLDERDEVICETYPE"
placeholder_device_name = "PLACEHOLDERDEVICENAME"
placeholder_device_loc = "PLACEHOLDERLOCATION"


# Create an empty list to store the updated rows
updated_rows = []
failed = []
for result_index, result_row in result_df.iterrows():
    try:

        template_df = pd.read_excel(f'templates/{result_row["Device_Type"]}.xlsx', sheet_name='Sheet1', header=None)
        template_df.columns = ["BLOCK TYPE", "TAG", "DESCRIPTION"]
        print(f"res_index = {result_index}, res_row = {result_row}")
        # Iterate over rows in the 'pumps_template' DataFrame
        for index, row in template_df.iterrows():
            # Replace placeholders with device type and device name
            updated_row = row.apply(lambda cell: cell.replace(placeholder_device_type, result_row['Device_Type']))
            updated_row = updated_row.apply(lambda cell: cell.replace(placeholder_device_loc, result_row['location']))
            updated_row = updated_row.apply(lambda cell: cell.replace(placeholder_device_name, result_row['Device_Name']))

            updated_rows.append(updated_row)
    except:
        failed.append(result_row)

failed_df = pd.DataFrame(failed)
failed_df.to_excel('failed.xlsx', index=False)
# Create a DataFrame from the updated rows
final_df = pd.DataFrame(updated_rows, columns=["BLOCK TYPE", "TAG", "DESCRIPTION"])

# Save the result to a new Excel file
final_df.to_excel('output2.xlsx', index=False)
