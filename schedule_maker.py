import pandas as pd
import datetime as dt
from ortools.sat.python import cp_model
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ------------------------------
# 1. Read input Excel
# ------------------------------
input_file = 'input.xlsx'
doctors_df = pd.read_excel(input_file, sheet_name='Doctors')

# Set the month and year for which the schedule is to be created
month_for_schedule = dt.datetime.now().month % 12 + 1
year_for_schedule = dt.datetime.now().year + 1 if month_for_schedule == 1 else dt.datetime.now().year

# e.g. first day = 2025-01-01, last day = 2025-01-31
first_day = dt.date(year_for_schedule, month_for_schedule, 1)
last_day = (first_day.replace(month=month_for_schedule % 12 + 1, day=1) - dt.timedelta(days=1))
dates = [first_day + dt.timedelta(days=i)
         for i in range((last_day - first_day).days + 1)]
print(first_day)
print(last_day)

# Identify weekdays/weekends
is_weekend = {d: d.weekday() >= 5 for d in dates}
# for date in is_weekend:
#     print(f"{date}: {'Weekend' if is_weekend[date] else 'Weekday'}")
    
# Convert unavailability days to datetime.date objects
unavailability = {}

for index, row in doctors_df.iterrows():
    if pd.notna(row['Unavailability']):
        unavailable_days = [
            dt.date(year_for_schedule, month_for_schedule, int(day))
            for day in row['Unavailability'].split(',')
        ]
    else:
        unavailable_days = []

    # Update DataFrame
    doctors_df.at[index, 'Unavailability'] = unavailable_days
    # Build dictionary
    unavailability[row['Doctor']] = set(unavailable_days)

doctors_list = doctors_df['Doctor'].tolist()
# print(doctors_list)
# print(unavailability)

# print(doctors_df.loc[doctors_df['Doctor'] == 'Μαρία'])

# ------------------------------
# 2. Build the model
# ------------------------------
model = cp_model.CpModel()

x = {}
for i, day in enumerate(dates):
    # print(f"Day {i+1} ({day}): {'Weekend' if is_weekend[day] else 'Weekday'}")
    for doc in doctors_list:
        if doc in unavailability and day in unavailability[doc]:
            continue
        
        x[(i, doc)] = model.NewBoolVar(f"x_{i}_{doc}")
        
# for i in x:
#     print(i, x[i])
    
# Exactly one doctor per day
for i in range(len(dates)):
    model.Add(sum(x.get((i,doc), 0) for doc in doctors_list) == 1)

# Equal total duties per doctor
total_days = len(dates)
min_days = total_days // len(doctors_list)
max_days = min_days if total_days % len(doctors_list) == 0 else min_days + 1
for doc in doctors_list:
    model.Add(sum(x.get((i,doc), 0) for i in range(len(dates))) >= min_days)
    model.Add(sum(x.get((i,doc), 0) for i in range(len(dates))) <= max_days)
    
# No consecutive duties
for i in range(len(dates)-1):
    for doc in doctors_list:
        model.Add(x.get((i,doc), 0) + x.get((i+1,doc), 0) <= 1)
        
# Equal number of weekend/weekday duties
total_weekends = sum(is_weekend[d] for d in dates)
min_wkend = total_weekends // len(doctors_list)
max_wkend = min_wkend if total_weekends % len(doctors_list) == 0 else min_wkend + 1

for doc in doctors_list:
    model.Add(
        sum(x.get((i,doc), 0) for i,d in enumerate(dates) if is_weekend[d]) >= min_wkend
    )
    model.Add(
        sum(x.get((i,doc), 0) for i,d in enumerate(dates) if is_weekend[d]) <= max_wkend
    )

penalty_terms = []
for doc in doctors_list:
    for i in range(len(dates) - 2):
        if (i,doc) in x and (i+2,doc) in x:
            b = model.NewBoolVar(f"everyother_{i}_{doc}")
            model.Add(x[(i,doc)] + x[(i+2,doc)] == 2).OnlyEnforceIf(b)
            model.Add(x[(i,doc)] + x[(i+2,doc)] != 2).OnlyEnforceIf(b.Not())
            penalty_terms.append(b)
            
# for doc in doctors_list:
#     for 

# Add to objective: minimize total "every other day" violations
model.Minimize(sum(penalty_terms))
# ------------------------------
# 3. Solve
# ------------------------------
solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 120
status = solver.Solve(model)

if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    raise RuntimeError("No feasible schedule found.")

# ------------------------------
# 4. Export to Excel
# ------------------------------
schedule = []
for i, day in enumerate(dates):
    for doc in doctors_list:
        if (x.get((i,doc)) is not None) and (solver.Value(x[(i,doc)]) == 1):
            schedule.append({"Date": day, "Assigned Doctor": doc})
            break

schedule_df = pd.DataFrame(schedule)
outfile = "monthly_schedule.xlsx"
schedule_df.to_excel(outfile, index=False)
print("Schedule created: monthly_schedule.xlsx")

# --- Apply styling with openpyxl ---
wb = load_workbook(outfile)
ws = wb.active

# Style definition
weekend_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
weekend_font = Font(color="FFFFFF", bold=True)

for row in range(2, ws.max_row + 1):  # skip header row
    date_cell = ws.cell(row=row, column=1)  # assuming Date is in col A
    date_value = pd.to_datetime(date_cell.value).date()
    if date_value.weekday() >= 5:  # 5 = Saturday, 6 = Sunday
        for col in range(1, ws.max_column + 1):
            print(f"Styling cell at row {row}, col {col} for weekend")
            cell = ws.cell(row=row, column=col)
            cell.fill = weekend_fill
            cell.font = weekend_font

wb.save(outfile)


# print number of duties per doctor
for doc in doctors_list:
    duties = sum(solver.Value(x.get((i,doc), 0)) for i in range(len(dates)))
    print(f"{doc}: {duties} duties")
    
# print number of weekend duties per doctor
for doc in doctors_list:
    wend_duties = sum(solver.Value(x.get((i,doc), 0)) for i,d in enumerate(dates) if is_weekend[d])
    print(f"{doc}: {wend_duties} weekend duties")
    
    
# print days of week that each doctor is assigned
for doc in doctors_list:
    assigned_days = [dates[i].strftime("%d %a") for i in range(len(dates)) if solver.Value(x.get((i,doc), 0)) == 1]
    print(f"{doc}: assigned on {', '.join(assigned_days)}")
    
    