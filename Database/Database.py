import pandas as pd
import sqlite3
#Only thing user needs to edit
excel_file_path = r"C:\Users\tanto\Desktop\Forecast\Final_Updated_Net-PEAK19.xlsx"
database_path = r"C:\Users\tanto\Desktop\Forecast\data\NetPeak2019.db"

def convert_month_year(sheet_name):
    month_name_map = {
        'มค': 'January', 'กพ': 'February', 'มีค': 'March',
        'เมย': 'April', 'พค': 'May', 'มิย': 'June',
        'กค': 'July', 'สค': 'August', 'กย': 'September',
        'ตค': 'October', 'พย': 'November', 'ธค': 'December'
    }
    month_thai = sheet_name.split('.')[0]
    year_thai = int(sheet_name.split('.')[1])
    
    month_name = month_name_map.get(month_thai, month_thai)
    year = 2500 + (year_thai) - 543  # Correct conversion from BE to CE
    
    return month_name, year

def clean_and_transform_data_with_holiday_notes(sheet_data, sheet_name):
    day_types = sheet_data.iloc[0, 1:].values
    holidays = sheet_data.iloc[1, 1:].values
    clean_data = sheet_data.drop([0, 1])
    clean_data.reset_index(drop=True, inplace=True)
    new_column_names = ["Time"] + [str(i) for i in range(1, len(clean_data.columns))]
    clean_data.columns = new_column_names
    
    # Convert month name to English and add a Year column
    month_name, year = convert_month_year(sheet_name)
    clean_data['Month'] = month_name
    clean_data['Year'] = year
    
    # Melt the DataFrame to long format
    formatted_data_expanded = pd.melt(clean_data, id_vars=['Time', 'Month', 'Year'], 
                                      var_name='Day', value_name='Value')
    formatted_data_expanded['Day_Type'] = formatted_data_expanded['Day'].apply(lambda x: day_types[int(x)-1])
    formatted_data_expanded['Holiday'] = formatted_data_expanded['Day'].apply(lambda x: 1 if pd.notna(holidays[int(x)-1]) else 0)
    formatted_data_expanded['Note'] = formatted_data_expanded['Day'].apply(lambda x: holidays[int(x)-1] if pd.notna(holidays[int(x)-1]) else '')
    
    return formatted_data_expanded

def add_daily_extremes(data):
    # Identify rows where 'Time' indicates max and min values
    max_rows = data[data['Time'] == 'สูงสุดของวัน']
    min_rows = data[data['Time'] == 'ต่ำสุดของวัน']
    
    # Create dictionaries to map Day and Month-Year to max and min values
    max_values = {(row['Day'], row['Month'], row['Year']): row['Value'] for index, row in max_rows.iterrows()}
    min_values = {(row['Day'], row['Month'], row['Year']): row['Value'] for index, row in min_rows.iterrows()}
    
    # Add new columns for daily max and min values
    data['Daily_Max'] = data.apply(lambda row: max_values.get((row['Day'], row['Month'], row['Year']), None), axis=1)
    data['Daily_Min'] = data.apply(lambda row: min_values.get((row['Day'], row['Month'], row['Year']), None), axis=1)
    
    # Remove rows where 'Time' is 'สูงสุดของวัน' or 'ต่ำสุดของวัน'
    data = data[~data['Time'].isin(['สูงสุดของวัน', 'ต่ำสุดของวัน'])]
    
    return data
# Load the Excel data
excel_data = pd.ExcelFile(excel_file_path)
sheet_names = excel_data.sheet_names

# Initialize an empty DataFrame to hold all the data
data_transformed_with_notes = pd.DataFrame()

# Process each sheet and append the cleaned data to the DataFrame
for sheet in sheet_names:
    sheet_data = pd.read_excel(excel_data, sheet_name=sheet)
    cleaned_data = clean_and_transform_data_with_holiday_notes(sheet_data, sheet)
    data_transformed_with_notes = pd.concat([data_transformed_with_notes, cleaned_data], ignore_index=True)

# Drop rows where 'Time' column has specific values
values_to_drop = [
    "ทุกภาค จากกรมอุตุฯ",
    "พลังงานไฟฟ้า/วัน(รวม Pump)",
    "พลังงานไฟฟ้า/วัน",
    "Day Peak",
    "Time",
    "Evening Peak"
]
data_transformed_with_notes_filtered = data_transformed_with_notes[
    ~data_transformed_with_notes['Time'].isin(values_to_drop)
]

# Apply the function to add daily extremes
data_transformed_with_extremes = add_daily_extremes(data_transformed_with_notes_filtered)

# Save the data with column names back to SQLite database
conn = sqlite3.connect(database_path)
cursor = conn.cursor()

# Create table with column names
query = """
CREATE TABLE IF NOT EXISTS Net_PEAK19 (
    Year INTEGER,
    Month TEXT,
    Time TEXT,
    Day INTEGER,
    Day_Type TEXT,
    Holiday INTEGER,
    Note TEXT,
    Value REAL,
    Daily_Max REAL,
    Daily_Min REAL
)
"""
cursor.execute(query)
conn.commit()

# Write the modified data to the table
data_transformed_with_extremes.to_sql('Net_PEAK19', conn, if_exists='replace', index=False)

# Close the connection
conn.close()
print("Completed")
