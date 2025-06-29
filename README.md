# Excel to iCalendar Converter
This Python script reads course Brandeis University schedule data from an Workday Excel file and generates a .ics calendar file compatible with applications like Google Calendar, Apple Calendar, and Outlook. Each row in the spreadsheet is converted into a recurring calendar event based on the course’s days, time, and location.

## 🚀 Features

- Reads course schedule data from Excel using `pandas`.
- Handles multiple meeting patterns per course.
- Automatically assigns events to the correct weekdays.
- Adds recurrence rules for weekly meetings until the course end date.
- Exports to a valid `.ics` iCalendar file.

  
## Requirements
```bash
pip install pandas openpyxl icalendar pytz
