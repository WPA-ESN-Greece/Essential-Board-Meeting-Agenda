// Agenda Sheet Template
const AGENDA_TEMPLATE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('#No | Date')

// Meetings Notes Google Document Template
const NOTES_TEMPLATE_DOC_URL = "https://docs.google.com/document/d/1yZNtfD299o0RZ4EDJDsBYzwB3hXGOmHvdQ2TtJGThpw/edit"


// Cell Ranges
const DAY_OF_THE_WEEK_CELL = "C3"
const SHEET_DATE_CELL = "C4"
const START_TIME_CELL = "C5"
const END_TIME_CELL = "C6"
const MEETIING_NUMBER_CELL = "C7"
const MEETING_URL_CELL = "C8"
const MEETING_NOTES_LINK_CELL = "C9"
const MEETING_NOTES_FOLDER_LINK_CELL = "G3"
const MEETING_CALENDAR_LINK_CELL = "G4"

var DAY_OF_THE_WEEK = String(AGENDA_TEMPLATE_SHEET.getRange(DAY_OF_THE_WEEK_CELL).getValue())


// Meetings Event Details
const ActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
const MEETIING_NAME = ActiveSpreadsheet.getName().split(" |",1)

  // Start Time
  const START_TIME = AGENDA_TEMPLATE_SHEET.getRange("C5").getValue()
    const START_TIME_HOURS = START_TIME.split(":",1)
    const START_TIME_MINUTES = START_TIME.split(":",2).slice(1,2)

  // End Time
  const END_TIME = AGENDA_TEMPLATE_SHEET.getRange("C6").getValue()
    const END_TIME_HOURS = END_TIME.split(":",1)
    const END_TIME_MINUTES = END_TIME.split(":",2).slice(1,2)

// Day of the week
const DATE_FORMAT = "dd/MM/yy"
const TIMEZONE = Session.getScriptTimeZone()

var DAY_OF_THE_WEEK = String(AGENDA_TEMPLATE_SHEET.getRange(DAY_OF_THE_WEEK_CELL).getValue())



// Event Details
const EVENT_DESCRIPTION = "Yet another Meeting..."
const EVENT_LOCATION = "📞 Google Meet"

  // Guests
  const EVENT_GUESTS = ["inikolarakis+esntest@gmail.com"]
  //Example: EVENT_GUESTS = ["board@esnsection.org"]

// Calendar ID to create the Event
const CALENDAR_ID = "c_899b9a4fd23838c5ec1ed73267826602b4100d7eaf4cc2d20ec2afeb18271706@group.calendar.google.com"
const CALENDAR_URL = "https://calendar.google.com/calendar/embed?src="+CALENDAR_ID+"%40group.calendar.google.com&ctz=Europe%2FAthens"
