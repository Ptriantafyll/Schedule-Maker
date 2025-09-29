import pandas as pd
import datetime as dt
from ortools.sat.python import cp_model

input_file = 'input.xlsx'
doctors_df = pd.read_excel(input_file, sheet_name='Doctors')

doctors_list = doctors_df['Doctor'].tolist()
print(doctors_list)

# Set the month and year for which the schedule is to be created
month_for_schedule = dt.datetime.now().month % 12 + 1
year_for_schedule = dt.datetime.now().year + 1 if month_for_schedule == 1 else dt.datetime.now().year

first_day = dt.date(year_for_schedule, month_for_schedule, 1)
last_day = (first_day.replace(month=month_for_schedule % 12 + 1, day=1) - dt.timedelta(days=1))
dates = [first_day + dt.timedelta(days=i)
         for i in range((last_day - first_day).days + 1)]
print(first_day)
print(last_day)

# Identify weekdays/weekends
is_weekend = {d: d.weekday() >= 5 for d in dates}
for date in is_weekend:
    print(f"{date}: {'Weekend' if is_weekend[date] else 'Weekday'}")
    
# Convert unavailability days to datetime.date objects
for index, row in doctors_df.iterrows():
    if pd.notna(row['Unavailability']):
        unavailable_days = [dt.date(year_for_schedule, month_for_schedule, int(day)) for day in row['Unavailability'].split(',')]
        doctors_df.at[index, 'Unavailability'] = unavailable_days
    else:
        doctors_df.at[index, 'Unavailability'] = []
        
print(doctors_df.head())
# print(doctors_df.loc[doctors_df['Doctor'] == 'Μαρία'])