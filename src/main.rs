use calamine::{open_workbook_auto, Data, Reader};
use chrono::{Datelike, Duration, NaiveDate};
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
fn format_text(row: &[Data]) -> String {
    let calendar_entry = row[CALENDAR_ENTRY].to_string();
    url_encode(&calendar_entry)
}

fn format_details(row: &[Data]) -> String {
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
        get_date_contacted(row),     // Date Contacted
        row[NOTES]               // Notes
    );
    url_encode(&details)
}

fn get_date_contacted(row: &[Data]) -> String {
    let serial = match row[DATE_CONTACTED] {
        Data::Float(x) => x - 2.0,
        _ => return String::new(),
    };
    let start = NaiveDate::from_ymd_opt(1900, 1, 1).expect("Invalid date");
    let date_option = start.checked_add_signed(Duration::days(serial as i64));

    if let Some(date) = date_option {
        format!("{}/{}/{}", date.month(), date.day(), date.year())
    } else {
        String::new()
    }
}

fn format_dates(row: &[Data]) -> String {
    match &row[DUE_DATE] {
        Data::String(date_str) => {
            // Assuming the date format in Excel is "M/D/YYYY" or "MM/DD/YYYY"
            let date = NaiveDate::parse_from_str(date_str, "%m/%d/%Y")
                .or_else(|_| NaiveDate::parse_from_str(date_str, "%-m/%-d/%Y"))
                .unwrap_or_else(|_| chrono::Local::now().naive_local().date());

            format!(
                "{:04}{:02}{:02}T140000Z/{:04}{:02}{:02}T143000Z",
                date.year(),
                date.month(),
                date.day(),
                date.year(),
                date.month(),
                date.day()
            )
        },
        _ => {
            // If DUE_DATE is not a string, use today's date
            let date = chrono::Local::now().naive_local().date();
            format!(
                "{:04}{:02}{:02}T140000Z/{:04}{:02}{:02}T143000Z",
                date.year(),
                date.month(),
                date.day(),
                date.year(),
                date.month(),
                date.day()
            )
        }
    }
}

fn format_location(row: &[Data]) -> String {
    let address = &row[ADDRESS].to_string();
    url_encode(address)
}

fn generate_calendar_url(text: &str, details: &str, dates: &str, location: &str) -> String {
    format!(
        "https://calendar.google.com/calendar/u/0/r/eventedit?text={}&details={}&dates={}&location={}&ctz=America/Phoenix",
        text, details, dates, location
    )
}

fn url_encode(text: &str) -> String {
    text.replace(" ", "+")
        .replace("&", "%26")
        .replace("#", "%23")
        .replace("nan", "")
}
