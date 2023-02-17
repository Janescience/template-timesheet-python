import openpyxl
from datetime import date,datetime, timedelta
import calendar
import json

# Create a new Excel workbook
workbook = openpyxl.Workbook()

# Get the current month
today = date.today()
month = today.strftime("%m")
year = today.strftime("%Y")

# Load the holidays data from a JSON file
with open(f"./holidays/{year}.json", "r",encoding="utf-8") as f:
    holidays = json.load(f)

# Convert the date strings in the holiday data to date objects
holiday_dates = [datetime.strptime(year + "-" + date, "%Y-%m-%d").date() for date in holidays.keys()]

# Add a new sheet with the name of the current month
sheet = workbook.active
sheet.title = "Timesheet_"+month+year

# Define the headers
headers = ["Day", "Time In", "Time Out", "OT Start", "OT Finish", "Manager Approve", "Detail", "Remark"]

# Add the headers to the first row
for i, header in enumerate(headers):
    sheet.cell(row=1, column=i+1, value=header)

# Get the number of days in the current month
num_days = calendar.monthrange(today.year, today.month)[1]

current_day = today.replace(day=1)
for i in range(num_days):
    date_to_check = datetime(year=2023,month=2,day=i+1).date()
    
    weekday = date_to_check.weekday()

    sheet.cell(row=i+2, column=1, value=current_day.strftime("%d"))
    if date_to_check in holiday_dates:
        holiday_description = holidays[date_to_check.strftime("%Y-%m-%d")]
        sheet.cell(row=i+2, column=2, value=holiday_description)
    else:
        if weekday == 5:
            sheet.cell(row=i+2, column=2, value="--------------------------------------------------------------------------------------- วันเสาร์  -------------------------------------------------------------------------------------")
        elif weekday == 6:
            sheet.cell(row=i+2, column=2, value="-------------------------------------------------------------------------------------- วันอาทิตย์  -----------------------------------------------------------------------------------")
        else:
            sheet.cell(row=i+2, column=2, value="08:30")
            sheet.cell(row=i+2, column=3, value="17:30")
            
    current_day = current_day + timedelta(days=1)

# Save the workbook to a file
workbook.save(f"Timesheet_{month}{year}.xlsx")

