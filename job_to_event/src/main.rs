use calamine::{open_workbook_auto, DataType, Reader};
use chrono::{Datelike, NaiveDate};
use std::env;
use webbrowser;

const DATE_CONTACTED: usize = 0;
const TASK: usize = 1;
const SAME_DAY: usize = 2;
const SCHEDULED: usize = 3;
const TRAVEL: usize = 4;
const JOB_NAME: usize = 5;
const ADDRESS: usize = 6;
const BILLING_FIRST_NAME: usize = 7;
const BILLING_LAST_NAME: usize = 8;
const BILLING_COMPANY: usize = 9;
const PHONE_NUMBER: usize = 10;
const ON_SITE_CONTACT: usize = 11;
const ON_SITE_NUMBER: usize = 12;
const GC_NAME: usize = 13;
const GC_NUMBER: usize = 14;
const PERMIT_NUMBER: usize = 15;
const PO_NUMBER: usize = 16;
const WATER_PURVEYOR: usize = 17;
const DUE_DATE: usize = 18;
const COORDINATION: usize = 19;
const PARTS: usize = 20;
const NOTES: usize = 21;
const CALENDAR_ENTRY: usize = 22;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 3 {
        println!("Usage: cargo run <file_path> <row_number>");
        std::process::exit(1);
    }

    let file_path = &args[1];
    let row_num: usize = args[2].parse().expect("Row number must be an integer");

    let mut workbook = open_workbook_auto(file_path).expect("Cannot open file");
    let sheet = workbook
        .worksheet_range_at(0)
        .expect("Cannot find sheet")
        .expect("Error reading sheet");

    if let Some(row) = sheet.rows().nth(row_num - 1) {
        let text = format_text(row);
        let details = format_details(row);
        let dates = format_dates(row);
        let location = format_location(row);
        let url = generate_calendar_url(&text, &details, &dates, &location);

        println!("Opening URL: {}", url);
        webbrowser::open(&url).expect("Failed to open URL");
    }
}

// the title
fn format_text(row: &[DataType]) -> String {
    let calendar_entry = row[CALENDAR_ENTRY].to_string(); // Assuming first column is Calendar Entry
    url_encode(&calendar_entry)
}

fn format_details(row: &[DataType]) -> String {
    // You would modify these indices based on your actual Excel structure
    let details = format!(
        "Job Name: {}%0ADue Date: {}%0ACust: {} {} {} {}%0ATask: {}%0ACoordination: {}%0AParts: {}%0AOnsite Contact: {} {}%0AGC Info: {} {}%0APermit #: {}%0AAddress: {}%0AWater Purveyor: {}%0APO #: {}%0ASame Day: {}%0AScheduled: {}%0ATravel: {}%0ADate Contacted: {}%0ANotes: {}%0A",
        row[JOB_NAME],           // Job Name
        row[DUE_DATE],           // Due Date
        row[BILLING_FIRST_NAME], // Cust
        row[BILLING_LAST_NAME],  // Cust
        row[BILLING_COMPANY],    // Cust
        row[PHONE_NUMBER],       // Cust
        row[TASK],               // Task
        row[COORDINATION],       // Coordination
        row[PARTS],              // Parts
        row[ON_SITE_CONTACT],    // Onsite Contact
        row[ON_SITE_NUMBER],     // Onsite Contact
        row[GC_NAME],            // GC Info
        row[GC_NUMBER],          // GC Info
        row[PERMIT_NUMBER],      // Permit #
        row[ADDRESS],            // Address
        row[WATER_PURVEYOR],     // Water Purveyor
        row[PO_NUMBER],          // PO #
        row[SAME_DAY],           // Same Day
        row[SCHEDULED],          // Scheduled
        row[TRAVEL],             // Travel
        row[0],     // Date Contacted
        row[NOTES]               // Notes
    );
    url_encode(&details)
}

fn format_dates(row: &[DataType]) -> String {
    let date_str = &row[DATE_CONTACTED].to_string(); // Assuming fourth column is Date Contacted
    if let Ok(date) = NaiveDate::parse_from_str(date_str, "%Y-%m-%d") {
        format!(
            "{:04}{:02}{:02}T000000Z/{:04}{:02}{:02}T010000Z",
            date.year(),
            date.month(),
            date.day(),
            date.year(),
            date.month(),
            date.day()
        )
    } else {
        String::new()
    }
}

fn format_location(row: &[DataType]) -> String {
    let address = &row[ADDRESS].to_string(); // Assuming fifth column is Address
    url_encode(address)
}

fn generate_calendar_url(text: &str, details: &str, dates: &str, location: &str) -> String {
    format!(
        "https://calendar.google.com/calendar/u/0/r/eventedit?text={}&details={}&dates={}&location={}&recur=RRULE:FREQ=WEEKLY;UNTIL=20251231T000000Z&ctz=America/Phoenix",
        text, details, dates, location
    )
}

fn url_encode(text: &str) -> String {
    text.replace(" ", "+")
        .replace("&", "%26")
        .replace("#", "%23")
        .replace("nan", "")
}
