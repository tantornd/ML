import pandas as pd
import sqlite3
from datetime import datetime, timedelta

db_path = r"C:\Users\tanto\Desktop\Forecast\data\NetPeak2021.db"
conn = sqlite3.connect(db_path)
time_query = "SELECT DISTINCT Time FROM Final ORDER BY Time;"
time_values = pd.read_sql(time_query, conn)

# Start date
start_date = datetime.strptime('2021-08-09', '%Y-%m-%d')

# Calculate the previous ... 
monday_dates = [(start_date - timedelta(weeks=i)) for i in range(8)]

formatted_monday_dates = [(date.strftime('%Y'), date.strftime('%B'), str(date.day)) for date in sorted(monday_dates)]
where_clause = " OR ".join(["(Year = ? AND Month = ? AND Day = ?)"] * len(formatted_monday_dates))
correct_date_query_dynamic_asc = f"""
SELECT *
FROM Final
WHERE ({where_clause})
AND Time = ?
"""

# Flatten the date parameters for the query
date_params = [param for date in formatted_monday_dates for param in date]

# Function to replace invalid characters in sheet names
def sanitize_sheet_name(sheet_name):
    return sheet_name.replace(':', '.').replace('/', '_').replace('\\', '_').replace('[', '_').replace(']', '_').replace('*', '_').replace('?', '_')

def classify_day_type(date, holidays):
    if date.strftime('%Y-%m-%d') in holidays:
        return 'Holiday'
    elif date.weekday() >= 5:
        return 'Weekend'
    else:
        return 'Weekday'
    
# Create a dictionary to store data for each Time value with corrected month names
data_by_time_dynamic_asc = {}
for time in time_values['Time']:
    params = date_params + [time]
    data_by_time_dynamic_asc[time] = pd.read_sql(correct_date_query_dynamic_asc, conn, params=params)

# Function to get the value for a specific date and the actual date as well
def get_value_and_date(date, time):
    date_str = [date.strftime('%Y'), date.strftime('%B'), str(int(date.strftime('%d')))]
    query = """
    SELECT Value, Holiday
    FROM Final
    WHERE Year = ? AND Month = ? AND Day = ? AND Time = ?
    """
    params = date_str + [time]
    result = pd.read_sql(query, conn, params=params)
    if not result.empty:
        value = result['Value'].values[0]
        actual_date = date.strftime('%Y/%m/%d')
        holidays = result['Holiday'].dropna().unique()
        day_type = classify_day_type(date, holidays)
        return value, actual_date, day_type
    return None, None, None

# Update the Excel file with the new values and reformat as specified
output_excel_path_dynamic_asc_reformatted = r"C:\Users\tanto\Desktop\Forecast\Book1.xlsx"
with pd.ExcelWriter(output_excel_path_dynamic_asc_reformatted, engine='openpyxl') as writer:
    for time in time_values['Time']:
        sheet_name = sanitize_sheet_name(time)
        data = data_by_time_dynamic_asc[time]
        for index, row in data.iterrows():
            current_date = datetime.strptime(f"{row['Year']}-{row['Month']}-{row['Day']}", '%Y-%B-%d')
            day_before = current_date - timedelta(days=1)
            week_before = current_date - timedelta(days=7)
            two_weeks_before = current_date - timedelta(days=14)
            
            data.at[index, 'X1'], data.at[index, 'X1_Date'], data.at[index, 'X1_Daytype'] = get_value_and_date(day_before, time)
            data.at[index, 'X2'], data.at[index, 'X2_Date'], data.at[index, 'X2_Daytype'] = get_value_and_date(week_before, time)
            data.at[index, 'X3'], data.at[index, 'X3_Date'], data.at[index, 'X3_Daytype'] = get_value_and_date(two_weeks_before, time)
            
        #Personal Formatting
        data = data.rename(columns={'Value': 'Actual Load'})
        data['Forecast Load'] = ''
        data['MAPE (%)'] = ''
        data.insert(0, 'Set', ['Train'] * (len(data) - 1) + ['Test'])
        columns_order = ['Set', 'Time', 'Month', 'Year', 'Day', 'Day_Type', 'Holiday', 'Note', 
                         'Daily_Max', 'Daily_Min', 'X1_Date', 'X2_Date', 'X3_Date','X1_Daytype','X2_Daytype','X3_Daytype', 'X1', 'X2', 'X3', 'Actual Load', 'Forecast Load','MAPE (%)']
        data = data[columns_order]
        
        data.to_excel(writer, sheet_name=sheet_name, index=False)

output_excel_path_dynamic_asc_reformatted
