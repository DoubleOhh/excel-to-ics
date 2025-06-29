import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta
import pytz

# Load the Excel file
file_path = ''  # Replace with your actual file path
df = pd.read_excel(file_path, engine='openpyxl')

# Create a new calendar
cal = Calendar()

# Timezone setup (assuming your times are in Eastern Time)
tz = pytz.timezone('America/New_York')


# Helper function to parse the meeting time
def parse_meeting_time(time_string):
    try:
        start_time, end_time = time_string.split(' - ')
        return datetime.strptime(start_time.strip(), '%I:%M %p'), datetime.strptime(end_time.strip(), '%I:%M %p')
    except Exception as e:
        print(f"Error parsing time: {e}")
        return None, None


# Helper function to get the first occurrence of a specific weekday
def get_first_weekday(start_date, weekday_index):
    # weekday_index: 0 = Monday, 1 = Tuesday, ..., 6 = Sunday
    while start_date.weekday() != weekday_index:
        start_date += timedelta(days=1)
    return start_date


# Day mapping
day_map = {'M': 0, 'T': 1, 'W': 2, 'R': 3, 'F': 4}

# Iterate through the rows of the DataFrame
for index, row in df.iterrows():
    course_name = row['Course Listing']
    meeting_patterns = row['Meeting Patterns']  # This could have multiple patterns separated by '\n'
    instructor = row['Instructor']
    start_date = pd.to_datetime(row['Start Date']).date()
    end_date = pd.to_datetime(row['End Date']).date()

    # Split the meeting patterns by newline to handle multiple patterns
    meeting_pattern_list = meeting_patterns.split('\n')

    # Process each meeting pattern separately
    for pattern in meeting_pattern_list:
        # Split each pattern into days, time, and location (assuming '|' separator)
        pattern_parts = pattern.split('|')

        if len(pattern_parts) < 2:
            print(f"Skipping invalid pattern: {pattern}")
            continue

        days = pattern_parts[0].strip()  # Days (e.g., M/W/F)
        time = pattern_parts[1].strip()  # Time (e.g., 12:00 PM - 1:30 PM)
        location = pattern_parts[2].strip() if len(pattern_parts) > 2 else 'No location provided'

        # Parse the start and end times
        start_time, end_time = parse_meeting_time(time)

        if start_time and end_time:
            # Create events for each day
            for day in days.split('/'):
                if day in day_map:
                    weekday_index = day_map[day]
                    first_day_date = get_first_weekday(start_date, weekday_index)  # Get first occurrence of this day

                    event = Event()
                    event.add('summary', course_name)
                    event.add('location', location)
                    event.add('description', f"Instructor: {instructor}")

                    # Set the start and end datetime for this event
                    event.add('dtstart', tz.localize(datetime.combine(first_day_date, start_time.time())))
                    event.add('dtend', tz.localize(datetime.combine(first_day_date, end_time.time())))

                    # Set the recurrence rule to repeat weekly until the end date
                    event.add('rrule',
                              {'freq': 'weekly', 'until': tz.localize(datetime.combine(end_date, end_time.time()))})

                    # Print event details to verify correctness
                    print(f"Event created for {course_name}")
                    print(f"  Days: {days}, Time: {time}, Location: {location}")
                    print(f"  Start Date: {first_day_date}, End Date: {end_date}")
                    print(f"  Start Time: {start_time}, End Time: {end_time}")
                    print(f"  Instructor: {instructor}")
                    print(f"  Recurrence until: {end_date}")
                    print("-" * 40)

                    # Add event to calendar
                    cal.add_component(event)

# Save the calendar to an .ics file
ics_file_path = 'courses_schedule_separated.ics'  # This will save the file in the current directory
with open(ics_file_path, 'wb') as f:
    f.write(cal.to_ical())

print(f"ICS file created successfully: {ics_file_path}")