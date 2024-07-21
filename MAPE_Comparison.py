import os
import pandas as pd

# Define the folder containing the Excel files
folder_path = r"C:\Users\tanto\Desktop\Forecast\2021(10)"

# Initialize a dictionary to store the MAPE values
mape_data = {}

# Iterate over each file in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        
        # Read the Excel file
        xls = pd.ExcelFile(file_path)
        
        # Iterate over each sheet in the Excel file
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Extract the relevant information
            day = df.iloc[11, 5]   # Column F Row 12
            month = df.iloc[11, 2] # Column C Row 12
            mape = df.iloc[11, 21] # Column V Row 12
            
            # Initialize the dictionary structure if not already present
            if month not in mape_data:
                mape_data[month] = {}
            if day not in mape_data[month]:
                mape_data[month][day] = []
            
            # Append the MAPE value to the list
            mape_data[month][day].append(mape)

# Calculate the average MAPE for each day of the week for every month
average_mape_data = {'Month': [], 'Day': [], 'Average MAPE': []}

for month in mape_data:
    for day in mape_data[month]:
        average_mape = sum(mape_data[month][day]) / len(mape_data[month][day])
        average_mape_data['Month'].append(month)
        average_mape_data['Day'].append(day)
        average_mape_data['Average MAPE'].append(average_mape)

# Convert the results to a DataFrame
average_mape_df = pd.DataFrame(average_mape_data)

# Write the results to a new Excel file
output_file_path = r"C:\Users\tanto\Desktop\Forecast\2021(10)MAPE\2021(10)MAPE.xlsx"
average_mape_df.to_excel(output_file_path, index=False)

print('Average MAPE calculation completed.')
