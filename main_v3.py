import pandas as pd
import os  # Import the 'os' module for file checking


# Function to process data
def process_data(input_file, output_file):
    df = pd.read_excel(input_file, sheet_name=0)

    # Extract device type, name, and location based on your criteria
    df['location'] = df['Tag Name'].str.split('-').str[1]
    df['Device_Type'] = df['Tag Name'].str.split('-').str[2]
    df['Device_Name'] = df['Tag Name'].str.split('-').str[3]

    # Drop empty cells (NaN values) in both 'Device_Type' and 'Device_Name' columns
    filtered_df = df.dropna(subset=['Device_Type', 'Device_Name', 'location'])

    # Remove duplicates based on both 'Device_Type' and 'Device_Name' columns
    filtered_df = filtered_df.drop_duplicates(subset=['Device_Type', 'Device_Name', 'location'])

    # Create a new DataFrame with extracted data
    result_df = filtered_df[['Device_Type', 'Device_Name', 'location']]

    # Save the result to a new Excel file
    result_df.to_excel(output_file, index=False)


# Function to replace placeholders in templates
def replace_placeholders(template_file, result_row, updated_rows, placeholders, failed):
    try:
        template_df = pd.read_excel(template_file, sheet_name='Sheet1', header=None)
        template_df.columns = ["BLOCK TYPE", "TAG", "DESCRIPTION"]

        for index, row in template_df.iterrows():
            updated_row = row.apply(lambda cell: cell
                                    .replace(placeholders['type'], result_row['Device_Type'])
                                    .replace(placeholders['loc'], result_row['location'])
                                    .replace(placeholders['name'], result_row['Device_Name'])
                                    )

            updated_rows.append(updated_row)



    except FileNotFoundError:
        print(f"Template file not found: {template_file}")
        failed.append(result_row)


def main():
    input_file = 'Auto Template.xlsx'
    staging_file = 'EQ_List.xlsx'
    output_file = 'rdy_device2.xlsx'
    placeholder_device_type = "PLACEHOLDERDEVICETYPE"
    placeholder_device_name = "PLACEHOLDERDEVICENAME"
    placeholder_device_loc = "PLACEHOLDERLOCATION"

    placeholders = {
        'type': placeholder_device_type,
        'name': placeholder_device_name,
        'loc': placeholder_device_loc
    }

    # Process data and save to 'rdy_device2.xlsx'
    process_data(input_file, staging_file)

    # Iterate over the result_df and replace placeholders in templates
    result_df = pd.read_excel(staging_file)
    failed = []
    updated_rows = []
    for result_index, result_row in result_df.iterrows():
        template_file = f'templates/{result_row["Device_Type"]}.xlsx'
        replace_placeholders(template_file, result_row, updated_rows, placeholders, failed)
        print(f"Processed {result_index}/{len(result_df) - 1}")

    # Save failed rows to 'failed.xlsx'
    failed_df = pd.DataFrame(failed)
    failed_df.to_excel('failed.xlsx', index=False)

    # Create a DataFrame from the updated rows
    final_df = pd.DataFrame(updated_rows, columns=["BLOCK TYPE", "TAG", "DESCRIPTION"])

    # Save the result to a new Excel file
    final_df.to_excel(output_file, index=False)


if __name__ == "__main__":
    main()
