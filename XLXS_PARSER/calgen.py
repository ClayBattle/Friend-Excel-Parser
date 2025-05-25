import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime
import pytz

# Load the spreadsheet
file_path = "raw_data.xlsx"  # Update with your file path
xls = pd.ExcelFile(file_path)
df = xls.parse("Sheet1")  # Update sheet name if needed

# Rename columns based on structure
df = df.iloc[6:].rename(columns={
    "Unnamed: 0": "Activity ID",
    "Unnamed: 1": "Activity Name 1",
    "Unnamed: 2": "Activity Name 2",
    "Unnamed: 3": "Activity Name 3"
})

# Drop rows without a valid Activity ID
df = df.dropna(subset=["Activity ID"])

# Filter rows where Activity ID starts with 'B9' or 'B3'
df = df[df["Activity ID"].astype(str).str.startswith(("B9", "B3"))]

# Remove 'A' from the end of every cell in the entire DataFrame
df = df.applymap(lambda x: str(x).rstrip(' A') if isinstance(x, str) else x)

# Initialize iCalendar file
cal = Calendar()

# Populate calendar events
for _, row in df.iterrows():
    event = Event()

    # Set event summary based on available activity names
    summary = ""
    if pd.notna(row["Activity Name 1"]):
        summary = row["Activity Name 1"]     
    elif  pd.notna(row["Activity Name 2"]):
        summary = row["Activity Name 2"]     
    else:
        summary = row["Activity Name 3"]

    summary = row[0] + " - " + summary

    event.add('summary', summary)
    
    # Dynamically determine start and finish dates
    start_col = None
    finish_col = None
    start_date = None
    finish_date = None
    
    for col in df.columns:
        try:
            # Only attempt to parse if the value looks like a valid date (not a number like 10)
            value = str(row[col]).strip()
            if value.isdigit():
                # Skip if the value is a numeric value like '10'
                continue
            
            # Attempt to parse each column as a date
            date_val = pd.to_datetime(value, errors='coerce', dayfirst=True)  # Added dayfirst=True for better flexibility
            
            if pd.notna(date_val):
                # First valid date found, assign it as start date
                if not start_date:
                    start_date = date_val
                    start_col = col
                    #print(f"Start date found: {start_date} in column: {start_col}")  # Debugging output
                # Second valid date found, assign it as finish date
                elif not finish_date:
                    finish_date = date_val
                    finish_col = col
                    #print(f"Finish date found: {finish_date} in column: {finish_col}")  # Debugging output
        except Exception as e:
            # Skip columns that can't be converted to a date
            print(f"Error parsing column '{col}': {e}")  # Debugging output
            continue
    
    # Log if no valid start or finish date was found
    if not start_date:
        print(f"Warning: No valid start date for event '{event.get('summary')}'")

    if not finish_date:
        print(f"Warning: No valid finish date for event '{event.get('summary')}'")
    
    # Add start date to event if valid
    if start_date:
        event.add('dtstart', start_date.replace(tzinfo=pytz.UTC))
    
    # Add finish date to event if valid
    if finish_date:
        event.add('dtend', finish_date.replace(tzinfo=pytz.UTC))
    
    # Add the event to the calendar
    cal.add_component(event)

# Save to an iCalendar file
ical_file_path = "extracted_events.ics"
with open(ical_file_path, 'wb') as f:
    f.write(cal.to_ical())

print(f"iCalendar file saved as {ical_file_path}")
