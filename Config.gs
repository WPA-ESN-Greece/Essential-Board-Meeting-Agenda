const ss = SpreadsheetApp.getActiveSpreadsheet()

// Agenda Sheet Template
const AGENDA_TEMPLATE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('#No | Date')

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
const MEETING_GUESTS_CELL = "C10"

// Meetings Notes Google Document Template
const NOTES_TEMPLATE_DOC_URL = "https://docs.google.com/document/d/1yZNtfD299o0RZ4EDJDsBYzwB3hXGOmHvdQ2TtJGThpw/edit" // The one in ESN Greece's Google Drive. 
let NOTES_TEMPLATE_DOC_ID = AGENDA_TEMPLATE_SHEET.getRange(MEETING_NOTES_LINK_CELL).getValue()

const DOCUMENTATION_URL = "https://docs.google.com/document/d/1lKIBvzzRSKa0mBPJoZYAi-dAaDVTO2xH4NJGr-bnxFA/edit?usp=sharing"


let DAY_OF_THE_WEEK = String(AGENDA_TEMPLATE_SHEET.getRange(DAY_OF_THE_WEEK_CELL).getValue())


// If this is "Time-driven Meeting Generation" the timeTriggered function will generate meeting essentials.
let Time_Driven_Meeting_Generation = AGENDA_TEMPLATE_SHEET.getRange("C2").getValue()


// Meetings Event Details
let ActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
let MEETIING_NAME = ActiveSpreadsheet.getName().split(" |",1)

  // Start Time
  let START_TIME = AGENDA_TEMPLATE_SHEET.getRange("C5").getValue()
    let START_TIME_HOURS = START_TIME.split(":",1)
    let START_TIME_MINUTES = START_TIME.split(":",2).slice(1,2)

  // End Time
  let END_TIME = AGENDA_TEMPLATE_SHEET.getRange("C6").getValue()
    let END_TIME_HOURS = END_TIME.split(":",1)
    let END_TIME_MINUTES = END_TIME.split(":",2).slice(1,2)

// Day of the week
const DATE_FORMAT = "dd/MM/yy"
let TIMEZONE = Session.getScriptTimeZone()




// Event Details
const EVENT_DESCRIPTION = "Yet another Meeting..."
const EVENT_LOCATION = "📞 Google Meet"

  // Guests email addresses. Also accepts Google Groups/ Mailing Lists.  
  let EVENT_GUESTS = ss.getRange(MEETING_GUESTS_CELL).getValue().split(",")
    // Example: EVENT_GUESTS = ["board@esnsection.org"] 

// Calendar ID to create the Event
const CALENDAR_ID = AGENDA_TEMPLATE_SHEET.getRange('G4').getValue()
const CALENDAR_URL = "https://calendar.google.com/calendar/u/0?cid="+CALENDAR_ID+"%40group.calendar.google.com&ctz=Europe%2FAthens"
