import pandas as pd

#USER DEPENDENT VARIABLES!!!!
file_path = r"C:\Users\tanto\Desktop\Copy of Net-PEAK19.xlsx"
output_file_path = r"C:\Users\tanto\Desktop\Forecast\Final_Updated_Net-PEAK19.xlsx"

excel_data = pd.ExcelFile(file_path)
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Write a placeholder sheet initially
    pd.DataFrame({"Placeholder": ["This will be removed if other sheets are processed."]}).to_excel(writer, sheet_name="Placeholder", index=False)
    placeholder_written = False

    for sheet_name in excel_data.sheet_names:
        try:
            # Load the data from the current sheet
            df = pd.read_excel(excel_data, sheet_name=sheet_name)
            
            # Check if the sheet has sufficient data for processing
            if df.shape[0] < 4 or df.shape[1] < 4:
                continue  # Skip sheets that don't have the minimum required structure

            # Create a new DataFrame for the cleaned data
            num_rows = 56  # Row A and times (48 timeslots + 3 additional rows + min/max)
            num_days = df.iloc[2, 3:].dropna(how='all').shape[0]
            cleaned_df = pd.DataFrame(index=range(num_rows), columns=range(1 + num_days))

            # Step 2: Write specified values in column A including 24:00
            times = pd.date_range(start="00:00", end="23:30", freq="30min").strftime("%H:%M").tolist() + ["24:00"]
            col_a = ["ID", "Day", "วันสำคัญ"] + times + ["สูงสุดของวัน", "ต่ำสุดของวัน"]
            cleaned_df.iloc[:len(col_a), 0] = col_a

            # Step 3: Write day numbers in row 1, starting from column B
            cleaned_df.iloc[0, 1:num_days+1] = range(1, num_days+1)

            # Step 4: Correct the day format in row 2
            numerical_values = pd.to_numeric(df.iloc[2, 3:3 + num_days], errors='coerce')
            if not numerical_values.isnull().all():
                days_corrected = pd.to_datetime(numerical_values, unit='D', origin='1899-12-30').dt.strftime('%A')
                cleaned_df.iloc[1, 1:num_days+1] = days_corrected
            else:
                cleaned_df.iloc[1, 1:num_days+1] = ['Invalid Date'] * num_days

            # Step 5: Copy and translate day names (row 3)
            day_names_thai = df.iloc[3, 3:3 + num_days].values
            day_translation = {
                'จันทร์': 'Monday',
                'อังคาร': 'Tuesday',
                'พุธ': 'Wednesday',
                'พฤหัสบดี': 'Thursday',
                'ศุกร์': 'Friday',
                'เสาร์': 'Saturday',
                'อาทิตย์': 'Sunday'
            }
            cleaned_df.iloc[2, 1:num_days+1] = [day_translation.get(day, day) for day in day_names_thai]

            # Step 6: Copy the correct rows of data (5th to 55th row) for the time intervals
            data_values = df.iloc[4:55, 3:3 + num_days]
            cleaned_df.iloc[3:3+len(data_values), 1:num_days+1] = data_values.values

            # Write the cleaned sheet to the new Excel file
            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            placeholder_written = True
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {e}")
            continue

    # Remove the placeholder if other sheets were successfully processed
    if placeholder_written:
        del writer.book["Placeholder"]

print(f"Cleaned and adjusted data saved to {output_file_path}")
