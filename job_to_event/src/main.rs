use calamine::{open_workbook_auto, Reader, DataType};
use chrono::{NaiveDate, Datelike};
use std::env;
use webbrowser;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 3 {
        println!("Usage: cargo run <file_path> <row_number>");
        std::process::exit(1);
    }

    let file_path = &args[1];
    let row_num: usize = args[2].parse().expect("Row number must be an integer");

    let mut workbook = open_workbook_auto(file_path).expect("Cannot open file");
    let sheet = workbook.worksheet_range_at(0).expect("Cannot find sheet").expect("Error reading sheet");

    if let Some(row) = sheet.rows().nth(row_num - 2) {
        let text = format_text(row);
        let details = format_details(row);
        let dates = format_dates(row);
        let location = format_location(row);
        let url = generate_calendar_url(&text, &details, &dates, &location);

        println!("Opening URL: {}", url);
        webbrowser::open(&url).expect("Failed to open URL");
    }
}

fn format_text(row: &[DataType]) -> String {
    let calendar_entry = row[0].to_string(); // Assuming first column is Calendar Entry
    url_encode(&calendar_entry)
}

fn format_details(row: &[DataType]) -> String {
    // You would modify these indices based on your actual Excel structure
    let details = format!("Job Name: {}%0ADue Date: {}%0A...", row[1], row[2]);
    url_encode(&details)
}

fn format_dates(row: &[DataType]) -> String {
    let date_str = &row[3].to_string(); // Assuming fourth column is Date Contacted
    if let Ok(date) = NaiveDate::parse_from_str(date_str, "%Y-%m-%d") {
        format!("{:04}{:02}{:02}T000000Z/{:04}{:02}{:02}T010000Z",
                date.year(), date.month(), date.day(),
                date.year(), date.month(), date.day())
    } else {
        String::new()
    }
}

fn format_location(row: &[DataType]) -> String {
    let address = &row[4].to_string(); // Assuming fifth column is Address
    url_encode(address)
}

fn generate_calendar_url(text: &str, details: &str, dates: &str, location: &str) -> String {
    format!(
        "https://calendar.google.com/calendar/u/0/r/eventedit?text={}&details={}&dates={}&location={}&recur=RRULE:FREQ=WEEKLY;UNTIL=20251231T000000Z&ctz=America/Phoenix",
        text, details, dates, location
    )
}

fn url_encode(text: &str) -> String {
    text.replace(" ", "+").replace("&", "%26").replace("#", "%23").replace("nan", "")
}
