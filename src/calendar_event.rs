use chrono::{Datelike};


pub struct CalendarEvent {
    title: String,
    description: String,
    date: String,
    location: String,
}

impl CalendarEvent {
    pub fn create(title: String, description: String, location: String) -> Self {
        Self {
            title: url_encode(&title),
            description: url_encode(&description),
            date: get_date(),
            location: url_encode(&location),
        }
    }

    pub fn generate_url(&self) -> String {
        format!(
            "https://calendar.google.com/calendar/u/0/r/eventedit?text={}&details={}&dates={}&location={}&ctz=America/Phoenix",
            self.title, self.description, self.date, self.location
        )
    }
}

fn url_encode(text: &str) -> String {
    text.replace(" ", "+")
        .replace("&", "%26")
        .replace("#", "%23")
        .replace("nan", "")
}

fn get_date() -> String {
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
