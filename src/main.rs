mod calendar_event;

use calamine::{open_workbook_auto, Data, Reader};
use std::env;
use webbrowser;
use crate::calendar_event::CalendarEvent;

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
        let event = CalendarEvent::create(
          row[CALENDAR_ENTRY].to_string(),
          format_details(row),
          row[ADDRESS].to_string()
        );

        let url = event.generate_url();

        println!("Opening URL: {}", url);
        webbrowser::open(&url).expect("Failed to open URL");
    }
}

fn format_details(row: &[Data]) -> String {
    format!(
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
        row[DATE_CONTACTED],     // Date Contacted
        row[NOTES]               // Notes
    )
}