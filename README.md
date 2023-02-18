# Template Timesheet with Python.
Create excel template timesheet in current month have Thai holidays with python.

## Require
- `python` or `python3`
- `pip` or `pip3`

## Installations
- `pip install openpyxl`

## Run
- `python timesheet.py [month] [year]`
  
> Month and year are optional. If not filled in, the system defaults to the current month and year.

## Example 

`python timesheet.py 2 2023`

File name : **Timesheet_022023.xlsx**

| Day	| Time In	| Time Out | OT Start	| OT Finish	| Manager Approve	 | Detail	               | Remark    |
| ----|---------| ---------| ---------|-----------|------------------|-----------------------|-----------|
| 01	| 08:30	  | 17:30		 |			
| 02	| 08:30	  | 17:30		 |			
| 03	| 08:30	  | 17:30		 |		
| 04	| วันเสาร์   ||||||						
| 05	| วันอาทิตย์    
