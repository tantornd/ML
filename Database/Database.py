import pandas as pd
import sqlite3

#User Dependent Variable!!!!!
excel_file_path = r"C:\Users\tanto\Desktop\Forecast\Final_Updated_Net-PEAK20.xlsx"
database_path = r"C:\Users\tanto\Desktop\Forecast\data\CombinedNetPeak2.db"

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
    year = 2500 + year_thai - 543
    return month_name, year

def clean_and_transform_data_with_holiday_notes(sheet_data, sheet_name):
    day_types = sheet_data.iloc[0, 1:].values
    holidays = sheet_data.iloc[1, 1:].values
    clean_data = sheet_data.drop([0, 1])
    clean_data.reset_index(drop=True, inplace=True)
    new_column_names = ["Time"] + [str(i) for i in range(1, len(clean_data.columns))]
    clean_data.columns = new_column_names
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
    max_rows = data[data['Time'] == 'สูงสุดของวัน']
    min_rows = data[data['Time'] == 'ต่ำสุดของวัน']
    max_values = {(row['Day'], row['Month'], row['Year']): row['Value'] for index, row in max_rows.iterrows()}
    min_values = {(row['Day'], row['Month'], row['Year']): row['Value'] for index, row in min_rows.iterrows()}
    data['Daily_Max'] = data.apply(lambda row: max_values.get((row['Day'], row['Month'], row['Year']), None), axis=1)
    data['Daily_Min'] = data.apply(lambda row: min_values.get((row['Day'], row['Month'], row['Year']), None), axis=1)
    data = data[~data['Time'].isin(['สูงสุดของวัน', 'ต่ำสุดของวัน'])]
    return data

excel_data = pd.ExcelFile(excel_file_path)
sheet_names = excel_data.sheet_names
data_transformed_with_notes = pd.DataFrame()
for sheet in sheet_names:
    sheet_data = pd.read_excel(excel_data, sheet_name=sheet)
    cleaned_data = clean_and_transform_data_with_holiday_notes(sheet_data, sheet)
    data_transformed_with_notes = pd.concat([data_transformed_with_notes, cleaned_data], ignore_index=True)
values_to_drop = [
    "ทุกภาค จากกรมอุตุฯ",
    "พลังงานไฟฟ้า/วัน(รวม Pump)",
    "พลังงานไฟฟ้า/วัน",
    "Day Peak",
    "Time",
    "Evening Peak",
    "Temp. ณ เวลา Peak",
    "Pump SNR + BB + LTK"
]
data_transformed_with_notes_filtered = data_transformed_with_notes[
    ~data_transformed_with_notes['Time'].isin(values_to_drop)
]
data_transformed_with_extremes = add_daily_extremes(data_transformed_with_notes_filtered)

conn = sqlite3.connect(database_path)
cursor = conn.cursor()

# Create table if it doesn't exist
query = """
CREATE TABLE IF NOT EXISTS Final (
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

# Append the new data to the existing table
data_transformed_with_extremes.to_sql('Final', conn, if_exists='append', index=False)

conn.close()
print("Completed")
