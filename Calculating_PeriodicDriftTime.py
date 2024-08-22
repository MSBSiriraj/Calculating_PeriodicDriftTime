import os
import pandas as pd
import numpy as np

## Specify parameters
ppm_error = 10
acq_time = 79.2
peak_num = 500
delta = 5

# Specify the folder containing the input text files (current directory)
input_folder = '.'

# Create a folder to store the output Excel files
output_folder = 'export_peaks_to_excel'
os.makedirs(output_folder, exist_ok=True)

# Iterate over files in the input folder
for file_name in os.listdir(input_folder):
    if file_name.endswith('.txt'):
        # Construct the full path for input and output files
        input_file_path = os.path.join(input_folder, file_name)
        output_file_path = os.path.join(output_folder, os.path.splitext(file_name)[0] + '.xlsx')

        # Read the text file into a Pandas DataFrame
        with open(input_file_path, 'r') as file:
            # Skip the first three lines (header and two lines)
            for _ in range(3):
                file.readline()

            # Read the remaining lines into a list
            lines = file.readlines()

        # Process the lines to extract the data
        data = []
        for line in lines[:-2]:  # Exclude the last two lines
            # Remove leading and trailing whitespace, split by tab
            values = line.strip().split('\t')
            # Convert numeric values to float
            values = [float(val) if val.replace('.', '', 1).isdigit() else val for val in values]
            data.append(values)

        # Create a DataFrame from the extracted data
        df = pd.DataFrame(data)

        # Shift the last row so that the fourth column becomes the first column
        df.iloc[-1] = df.iloc[-1].shift(-3)

        # Transpose the DataFrame
        df_transposed = df.transpose()

        # Convert all cells to float
        df_transposed = df_transposed.apply(pd.to_numeric, errors='coerce')

        # Remove the last two rows and any rows with NaN values
        df_transposed = df_transposed.iloc[:-2, :].dropna()

        # Save the modified and transposed DataFrame to an Excel file
        df_transposed.to_excel(output_file_path, index=False, header=False)

        print(f"Conversion from {file_name} to {output_file_path} completed successfully.")

os.chdir("export_peaks_to_excel")

# Input folder path and output folder path
input_folder_path = "."
output_folder_path = "./Converted_DriftTime"

# Create the output folder if it doesn't exist
os.makedirs(output_folder_path, exist_ok=True)

# List all Excel files in the input folder
excel_files = [file for file in os.listdir(input_folder_path) if file.endswith('.xlsx')]

for input_file_name in excel_files:
    # Generate output file path
    output_file_name = os.path.join(output_folder_path, input_file_name)

    # Step 1: Delete whatever is before and unto "separate"
    file_name_parts = input_file_name.split("separate")[-1]

    # Step 2: If the leftover from 1 contains "_", delete everything from "_" to the end
    if "_" in file_name_parts:
        file_name_parts = file_name_parts.split("_")[0]
    # If the leftover from 1 does not contain "_", delete ".xlsx"
    else:
        file_name_parts = os.path.splitext(file_name_parts)[0]

    # Step 3: If the leftover from 2 contains "p", replace "p" with "."
    file_name_parts = file_name_parts.replace("p", ".")

    # Step 4: Convert to float
    B_values = float(file_name_parts) if file_name_parts else None

    # Check if B value was successfully extracted
    if B_values is not None:
        # Read the Excel file
        df = pd.read_excel(os.path.join(input_folder_path, input_file_name), header=None)

        # Extract values from the second and third columns
        C_values = df.iloc[:, 1]

        # Calculate drift time in ms (A)
        df[1] = 10 + B_values + (acq_time / 200) * C_values

        # Write the output to a new Excel file
        df.to_excel(output_file_name, header=False, index=False)

        print(f"B = {B_values}")
        print(f"Conversion completed. Output saved to {output_file_name}")
    else:
        print(f"Unable to extract B value from the file name: {input_file_name}")

# Sort the peaks by intensity
file_to_open = [file for file in os.listdir(output_folder_path) if file.endswith('separate0p01.xlsx')][0]
df = pd.read_excel(os.path.join(output_folder_path, file_to_open), header=None)
df_sorted = df.sort_values(by=[2], ascending=False)


# Write the top peaks into a file called Top_peaks.xlsx at Converted_DriftTime folder
top_peaks = df_sorted.head(peak_num)

# Rename the columns
top_peaks.columns = ["m/z", "DT", "Intensity"]

top_peaks.to_excel(os.path.join(output_folder_path, 'Top_peaks.xlsx'), index=False)
print("Top peaks sorted by intensity have been written to Converted_DriftTime/Top_peaks.xlsx")

os.chdir("Converted_DriftTime")

import os
import pandas as pd
import numpy as np

# Set the Excel folder to the current working directory
excel_folder = os.getcwd()
print(excel_folder)

# Read the reference file
df_reference = pd.read_excel('../../Top_peaks_Human22CwithStandards.xlsx')

# Create the "matched_mz" folder if it doesn't exist
output_folder = os.path.join(excel_folder, 'matched_mz')
os.makedirs(output_folder, exist_ok=True)

# Create an empty DataFrame to store all drift time columns
all_drift_time_df = pd.DataFrame()

# Concatenate the 'm/z' column from df_reference to all_drift_time_df as the first column
all_drift_time_df['m/z'] = df_reference['m/z']

# Loop through each Excel file in the folder
for excel_file in os.listdir(excel_folder):
    if excel_file.endswith('.xlsx') and excel_file != 'Top_peaks.xlsx':
        # Read the Excel file without headers
        df_excel = pd.read_excel(os.path.join(excel_folder, excel_file), header=None)

        # Step 1: Delete whatever is before and unto "separate"
        file_name_parts = excel_file.split("separate")[-1]

        # Step 2: If the leftover from 1 contains "_", delete everything from "_" to the end
        if "_" in file_name_parts:
            file_name_parts = file_name_parts.split("_")[0]
        # If the leftover from 1 does not contain "_", delete ".xlsx"
        else:
            file_name_parts = os.path.splitext(file_name_parts)[0]

        # Step 3: If the leftover from 2 contains "p", replace "p" with "."
        file_name_parts = file_name_parts.replace("p", ".")

        # Step 4: Convert to float
        B_values = float(file_name_parts) if file_name_parts else None

        # Create a column header for the detected drift time
        drift_time_column = B_values
        print(drift_time_column)

        # Iterate over each row in the reference DataFrame
        for _, ref_row in df_reference.iterrows():
            reference_m_z = ref_row['m/z']

            # Calculate the ppm difference and filter rows in the Excel file
            ppm_diff = reference_m_z * ppm_error*1e-6
            filtered_rows = df_excel[
                (df_excel[0] >= (reference_m_z - ppm_diff)) &
                (df_excel[0] <= (reference_m_z + ppm_diff))
            ]

            # If there are matching rows, select the one with the highest intensity
            if not filtered_rows.empty:
                # Sort by intensity in descending order
                sorted_rows = filtered_rows.sort_values(by=2, ascending=False)
                # Take the first row and extract the drift time
                drift_time = sorted_rows[1].iloc[0]
            else:
                # If no match is found, set drift time to NaN
                drift_time = np.nan

            # Append the drift time to the results DataFrame in the corresponding column
            df_reference.loc[df_reference['m/z'] == reference_m_z, drift_time_column] = drift_time

        # Select only the desired columns for the output
        output_df = df_reference[['m/z', drift_time_column]]

        # Store the drift time column in the all_drift_time_df
        all_drift_time_df[drift_time_column] = df_reference[drift_time_column]

# Write all drift time columns to a single Excel file
all_drift_time_df.to_excel(os.path.join(excel_folder, 'Matched_mz/all_arrival_times.xlsx'), index=False)

# Write all drift time columns to a single Excel file
all_drift_time_df.to_excel(os.path.join(excel_folder, '../../all_arrival_times.xlsx'), index=False)

os.chdir("../../")

# Display a message when the processing is complete
print("Processing complete. Results saved in the 'matched_mz' folder. All drift times saved in 'all_arrival_times.xlsx'")

import pandas as pd

# Read input Excel file
df = pd.read_excel("all_arrival_times.xlsx")

# Create a new DataFrame for the output
output_df = df[['m/z']].copy()

# Iterate through each compound
for index, row in df.iterrows():
    compound_data = row[1:].dropna().tolist()  # Remove NaN values and extract drift time data
    compound_data.sort()  # Sort the drift time data

    dummy_list = []
    pass_number = 0

    for i in range(len(compound_data)):
        if i == 0:
            dummy_list.append(compound_data[i])
        else:
            if compound_data[i] - compound_data[i - 1] <= delta:
                dummy_list.append(compound_data[i])
            else:
                avg_value = sum(dummy_list) / len(dummy_list)
                output_df.at[index, pass_number] = avg_value
                dummy_list = [compound_data[i]]
                pass_number += 1

    # Handle the last set of drift times for the compound
    if len(dummy_list) > 5:
        avg_value = sum(dummy_list) / len(dummy_list)
    else:
        avg_value = 0
    output_df.at[index, pass_number] = avg_value

# Fill NaN values with 0 in the output DataFrame
output_df.fillna(0, inplace=True)

# Save the output DataFrame to an Excel file
output_df.to_excel("ArrivalTime_by_Pass.xlsx", index=False)

import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score  # Import r2_score function

# Load the data
df = pd.read_excel('ArrivalTime_by_Pass.xlsx')

# Remove columns 0 and 1
# df = df.drop(columns=[0], errors='ignore')

# Calculate periodic drift time, y-intercept, and R-square for each row
periodic_drift_times = []
y_intercepts = []
r_squared_values = []  # Initialize list to store R-square values
for index, row in df.iterrows():
    x = [int(col) for col in row.index[1:]]  # Pass numbers
    y = row.values[1:]  # Drift times
    # Remove zero values
    x_nonzero = [x_val for x_val, y_val in zip(x, y) if y_val != 0]
    y_nonzero = [y_val for y_val in y if y_val != 0]
    if len(x_nonzero) < 4:
        periodic_drift_times.append(None)
        y_intercepts.append(None)
        r_squared_values.append(None)  # Append None for rows where R-square couldn't be calculated
    else:
        # Fit linear regression
        model = LinearRegression()
        model.fit([[x_val] for x_val in x_nonzero], y_nonzero)
        # Calculate the slope (periodic drift time)
        periodic_drift_times.append(model.coef_[0])
        # Calculate the y-intercept
        y_intercepts.append(model.intercept_)
        # Calculate R-square
        r_squared_values.append(r2_score(y_nonzero, model.predict([[x_val] for x_val in x_nonzero])))

# Add the calculated periodic drift times, y-intercepts, and R-square values to the DataFrame
df['Periodic Drift Time'] = periodic_drift_times
df['Y-Intercept'] = y_intercepts
df['R-Square'] = r_squared_values

# # Drop rows where either periodic drift time, y-intercept, or R-square couldn't be calculated
# df.dropna(subset=['Periodic Drift Time', 'Y-Intercept', 'R-Square'], inplace=True)

# Save the result to a new Excel file with two sheets
with pd.ExcelWriter('periodic_dt.xlsx') as writer:
    df[['m/z', 'Periodic Drift Time', 'Y-Intercept', 'R-Square']].to_excel(writer, sheet_name='Sheet1', index=False)

