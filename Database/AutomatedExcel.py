import pandas as pd
import sqlite3
from datetime import datetime, timedelta

db_path = '/mnt/data/NetPeak2021.db'
conn = sqlite3.connect(db_path)
time_query = "SELECT DISTINCT Time FROM Final ORDER BY Time;"
time_values = pd.read_sql(time_query, conn)

# Start date
start_date = datetime.strptime('2021-08-09', '%Y-%m-%d')

# Calculate the previous x-1 days
x = 8
monday_dates = [(start_date - timedelta(weeks=i)).strftime('%Y-%B-%d').split('-') for i in range(x)]

where_clause = " OR ".join(["(Year = ? AND Month = ? AND Day = ?)"] * len(monday_dates))
correct_date_query_dynamic = f"""
SELECT *
FROM Final
WHERE ({where_clause})
AND Time = ?
ORDER BY Year DESC, Month DESC, Day DESC;
"""

# Flatten the date parameters for the query
date_params = [param for date in monday_dates for param in date]

# Function to replace invalid characters in sheet names
def sanitize_sheet_name(sheet_name):
    return sheet_name.replace(':', '.').replace('/', '_').replace('\\', '_').replace('[', '_').replace(']', '_').replace('*', '_').replace('?', '_')

# Create a dictionary to store data for each Time value with corrected month names
data_by_time_dynamic = {}
for time in time_values['Time']:
    params = date_params + [time]  # Combine date parameters with the current time
    data_by_time_dynamic[time] = pd.read_sql(correct_date_query_dynamic, conn, params=params)

# Save the corrected data to an Excel file with each sheet representing a Time value
output_excel_path_dynamic = '/mnt/data/Mondays_Data_By_Time_Dynamic.xlsx'
with pd.ExcelWriter(output_excel_path_dynamic) as writer:
    for time, data in data_by_time_dynamic.items():
        data.to_excel(writer, sheet_name=sanitize_sheet_name(time), index=False)

output_excel_path_dynamic
