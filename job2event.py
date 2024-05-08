import pandas as pd
import webbrowser
import sys

def load_data(file_path):
    data = pd.read_excel(file_path, engine='openpyxl')
    return data

# Title Section
def format_text(row):
    return f"{row['Calendar Entry']}".replace(' ', '+').replace('&', '%26')

# Comment Section
def format_details(row):
    details = [
        "Job Name: " + str(row['Job Name']), '%0A',
        "Due Date: " + str(row['Due Date']), '%0A',
        "Cust: " + f"{row['Billing First Name']} {row['Billing Last Name']} w {row['Billing Company']} @ {row['Phone #']}", '%0A',
        "Task: " + str(row['Task']), '%0A',
        "Coordination: " + str(row['Coordination']), '%0A',
        "Parts: " + str(row['Parts']), '%0A',
        "Onsite Contact: " + f"{row['On Site Contact']} @ {row['On Site #']}", '%0A',
        "GC Info: " + f"{row['GC Name']} @ {row['GC #']}", '%0A',
        "Permit #: " + str(row['Permit #']), '%0A',
        "Address: " + str(row['Address']), '%0A',
        "Water Purveyor: " + str(row['WaterPurveyor']), '%0A',
        "PO #: " + str(row['PO #']), '%0A',
        "Scheduled: " + str(row['Scheduled?']), '%0A',
        "Travel: " + str(row['Travel']), '%0A',
        "Date Contacted: " + str(row['Date Contacted']), '%0A',
        "Notes: " + str(row['Notes']), '%0A',
    ]
    return '+'.join(detail.replace(' ', '+').replace('&', '%26') for detail in details).replace('#', '%23').replace('nan', '')

# Date Section
def format_dates(row):
    date_formatted = pd.to_datetime(row['Date Contacted']).strftime('%Y%m%d')
    return f"{date_formatted}T000000Z/{date_formatted}T010000Z"

# Location Section
def format_location(row):
    return str(row['Address']).replace(' ', '+').replace('&', '%26')

# URL Assembler
def generate_calendar_url(text, details, dates, location):
    return (
        "https://calendar.google.com/calendar/u/0/r/eventedit?"
        f"text={text}&details={details}&dates={dates}&location={location}"
        "&recur=RRULE:FREQ=WEEKLY;UNTIL=20251231T000000Z&ctz=America/Phoenix"
    )

def generate_calendar_url_for_row(data, row_num):
    row = data.iloc[row_num - 2]
    text = format_text(row)
    details = format_details(row)
    dates = format_dates(row)
    location = format_location(row)
    return generate_calendar_url(text, details, dates, location)

def main():
    if len(sys.argv) < 3:
        print("Usage: python job2event.py <file_path> <row_number>")
        sys.exit(1)

    file_path = sys.argv[1]  # Get file path from command-line arguments
    row_num = int(sys.argv[2])  # Get row number from command-line arguments

    data = load_data(file_path)
    url = generate_calendar_url_for_row(data, row_num)

    webbrowser.open(url)

if __name__ == '__main__':
    main()
