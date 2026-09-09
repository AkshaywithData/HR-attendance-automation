# HR Attendance Automation  

## Project Overview

This project automates HR attendance processing by converting raw attendance data into Excel reports and dashboards. It performs data cleaning, attendance analysis, KPI calculations, and chart generation, helping HR teams monitor daily, monthly, and yearly attendance records.

## Features

- Supports both CSV and Excel (.xlsx) attendance files.
- Automatically maps employee IDs to employee names using an Employee Master file.
- Cleans and transforms raw attendance data into an analysis-ready format.
- Calculates employee working hours from punch-in and punch-out times.  
- Calculates overtime hours for each employee.
- Automatically marks employees as Present or Absent.
- Computes key HR metrics.
- Generates comprehensive reports.
- Flags attendance exceptions.
- Creates automated charts.
- Builds a professional Excel dashboard using a reusable dashboard template.
- Automatically embeds generated charts into the Excel dashboard.
- Saves the cleaned attendance dataset for future analysis.

## Technologies Used

- Python
- Pandas
- NumPy
- OpenPyXL
- Datetime
- OS
- Matplotlib
- Shutil
- Glob

## How to Run

### 1. Install Required Libraries

```bash
pip install -r requirements.txt
```

### 2. Prepare the Input Files

Place your files in the following folders:

```
Data/
├── Source Files/
│   ├── Attendance file/
│   │   └── attendance_2023_2024.csv (or .xlsx)
│   │
│   └── Employees file/
│       └── Employee_Master.xlsx
```
```
 Generated Reports/
    └── HR_Attendance_Report.xlsx
```

  Place Employee_Master.xlsx in the Employees file folder and place HR_Attendance_Report.xlsx in the Generated Reports folder. The workbook should contain the required dashboard template.

### 3. Run the Project

```bash
python HR_automation.py
```

## Project Structure

```text
HR_Attendance_Automation/
│
├── HR_automation.py
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
│
├── Data/
│   ├── Source Files/
│   │   ├──Attendance file 
│   │   │         └── attendance_2023_2024.csv
│   │   └──Employees file
│   │             └── Employee_Master.xlsx
│   │
│   ├── Clean_Files/  
│   │
│   └── Archive/
│
├── Generated Reports/
│   ├── HR_Attendance_Report.xlsx
│   │       ├── Employee_Summary
│   │       ├── Daily_Summary
│   │       ├── Exception_Report
│   │       ├── Overtime_summary
│   │       ├── Dashboard
│   │       └── Dashboard_Template
│   │ 
│   └── Charts/
│       ├── attendance_trend.png
│       ├── monthly_ot.png
│       ├── top_late_employees.png
│       └── top_ot_employees.png
│       
│
└── Templates/
  └── Dashboard_Template.xlsx
```

## Output

After execution, the project generates:

- Reads attendance files
- Cleans and transforms the data
- Maps employee names
- Calculates attendance KPIs
- Generates employee and daily summary reports
- Generates overtime summaries
- Creates automated charts
- Builds an Excel dashboard
- Saves the Excel final report
- Archives processed attendance files

## Future Improvement

- Add email automation for sending attendance reports
- Add database integration for storing attendance data
- Add API integration for connecting with HR systems
- Add automated scheduling for daily or monthly processing
- Add advanced dashboard filtering and analytics


## License

This project is licensed under the MIT License.


## Author

**Akshay Gawand**