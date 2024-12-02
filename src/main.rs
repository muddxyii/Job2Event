use calamine::{open_workbook_auto, Data, Reader};
use std::env;

use actix_web::{get, web, App, HttpResponse, HttpServer, Responder};
use actix_files as fs;
use webbrowser;

use serde::Serialize;

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



#[derive(Serialize)]
pub struct CalendarEvent {
    pub title: String,
    pub description: String,
    pub address: String,
}

#[get("/api/event-data")]
async fn get_event_data() -> impl Responder {
    let args: Vec<String> = env::args().collect();
    let event = if args.len() >= 3 {
        let file_path = &args[1];
        let row_num: usize = args[2].parse().unwrap_or(1);

        let mut workbook = open_workbook_auto(file_path).unwrap_or_else(|_| panic!("Cannot open file"));
        let sheet = workbook.worksheet_range_at(0).unwrap().unwrap();

        if let Some(row) = sheet.rows().nth(row_num - 1) {
            CalendarEvent {
                title: row[CALENDAR_ENTRY].to_string(),
                description: format_details(row),
                address: row[ADDRESS].to_string()
            }
        } else {
            CalendarEvent {
                title: "Default Event".to_string(),
                description: "No data found".to_string(),
                address: "".to_string()
            }
        }
    } else {
        CalendarEvent {
            title: "Default Event".to_string(),
            description: "No file provided".to_string(),
            address: "".to_string()
        }
    };

    web::Json(event)
}

#[get("/")]
async fn index() -> impl Responder {
    HttpResponse::Ok()
        .content_type("text/html")
        .body(include_str!("../static/index.html"))
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    let server = HttpServer::new(|| {
        App::new()
            .service(index)
            .service(get_event_data)
            .service(fs::Files::new("/static", "static").show_files_listing())
    })
        .bind("127.0.0.1:8080")?;

    webbrowser::open("http://localhost:8080/").expect("Failed to open URL");

    server.run().await
}

fn format_details(row: &[Data]) -> String {
    format!(
        "Job Name: {}\nDue Date: {}\nCust: {} {} {} {}\nTask: {}\nCoordination: {}\nParts: {}\nOnsite Contact: {} {}\nGC Info: {} {}\nPermit #: {}\nAddress: {}\nWater Purveyor: {}\nPO #: {}\nSame Day: {}\nScheduled: {}\nTravel: {}\nDate Contacted: {}\nNotes: {}\n",
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