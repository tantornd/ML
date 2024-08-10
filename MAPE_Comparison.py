import os
import pandas as pd
#User Dependent Variables!!!!!
folder_path = r"C:\Users\tanto\Desktop\Forecast\2021(4)"
output_file_path = r"C:\Users\tanto\Desktop\Forecast\MAPEComparison\2021(4)MAPE - Copy.xlsx"
mape_data = {}

def add_bridge_holiday(mape_data, month, day, mape):
    if day == 'Tuesday':
        bridge_day = 'Monday'
    elif day == 'Thursday':
        bridge_day = 'Friday'
    else:
        return
    
    if bridge_day not in mape_data[month]['bridge_holidays']:
        mape_data[month]['bridge_holidays'][bridge_day] = []
    mape_data[month]['bridge_holidays'][bridge_day].append(mape)

# Iterate over each file in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        xls = pd.ExcelFile(file_path)
        
        # Iterate over each sheet in the Excel file
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            day = df.iloc[-1, 5]
            month = df.iloc[-1, 2]
            mape = df.iloc[-1, 21]
            is_holiday = df.iloc[-1, 6] == 1  
            
            # Initialize the dictionary structure if not already present
            if month not in mape_data:
                mape_data[month] = {'days': {}, 'holidays': [], 'bridge_holidays': {}}
            if day not in mape_data[month]['days']:
                mape_data[month]['days'][day] = []
            
            # Append the MAPE value to the list
            if is_holiday:
                mape_data[month]['holidays'].append((day, mape))
                add_bridge_holiday(mape_data, month, day, mape)
            else:
                mape_data[month]['days'][day].append(mape)

# Prepare the data for writing to Excel
formatted_data = {
    'Month': [],
    'Monday': [],
    'Tuesday': [],
    'Wednesday': [],
    'Thursday': [],
    'Friday': [],
    'Saturday': [],
    'Sunday': [],
    'Holiday': [],
    'Bridge Holiday': []
}

days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

for month in mape_data:
    formatted_data['Month'].append(month)
    for day in days_of_week:
        if day in mape_data[month]['days']:
            avg_mape = sum(mape_data[month]['days'][day]) / len(mape_data[month]['days'][day])
            formatted_data[day].append(avg_mape)
        else:
            formatted_data[day].append(None)
    
    if mape_data[month]['holidays']:
        avg_holiday_mape = sum(mape[1] for mape in mape_data[month]['holidays']) / len(mape_data[month]['holidays'])
        formatted_data['Holiday'].append(avg_holiday_mape)
    else:
        formatted_data['Holiday'].append(None)
    
    if mape_data[month]['bridge_holidays']:
        avg_bridge_mape = []
        for bridge_day in mape_data[month]['bridge_holidays']:
            avg_bridge_mape.extend(mape_data[month]['bridge_holidays'][bridge_day])
        if avg_bridge_mape:
            formatted_data['Bridge Holiday'].append(sum(avg_bridge_mape) / len(avg_bridge_mape))
        else:
            formatted_data['Bridge Holiday'].append(None)
    else:
        formatted_data['Bridge Holiday'].append(None)

# Convert the results to a DataFrame
average_mape_df = pd.DataFrame(formatted_data)

# Write the results to a new Excel file
average_mape_df.to_excel(output_file_path, index=False)

print('Average MAPE calculation and formatting completed.')
