function newMeetingEssentials()
{
  var activeSheet = ActiveSpreadsheet.getActiveSheet()
  
  var NOTES_TEMPLATE_DOC_ID = extractDocumentIdFromUrl(NOTES_TEMPLATE_DOC_URL)
  var NOTES_TEMPLATE_DOC = DriveApp.getFileById(NOTES_TEMPLATE_DOC_ID)

  // Start Time
  const START_TIME = AGENDA_TEMPLATE_SHEET.getRange(START_TIME_CELL).getValue()
    const START_TIME_HOURS = START_TIME.split(":",1)
    const START_TIME_MINUTES = START_TIME.split(":",2).slice(1,2)

  // End Time
  const END_TIME = AGENDA_TEMPLATE_SHEET.getRange(END_TIME_CELL).getValue()
    const END_TIME_HOURS = END_TIME.split(":",1)
    const END_TIME_MINUTES = END_TIME.split(":",2).slice(1,2)

  // Meeting Variables
    var meetingDateFormated
    var meetingNumber = Number(activeSheet.getRange(MEETIING_NUMBER_CELL).getValue())
    var meetingDate 
    var meetingTitle = String(MEETIING_NAME)
    var meetingAgendaURL = ""
    var meetingNotesDoc
    var meetingNotesDocURL = ""
    var meetingMeetURL = ""
    var meetingStartTime = new Date(null,null,null,START_TIME_HOURS,START_TIME_MINUTES)
    var meetingEndTime = new Date(null,null,null,END_TIME_HOURS,END_TIME_MINUTES)
  
  var sheetDate = new Date(activeSheet.getRange(SHEET_DATE_CELL).getValue())

  //Agenda Parent folder
  var destinationFolderID = DriveApp.getFileById(ActiveSpreadsheet.getId()).getParents().next().getId()

  // Calculates the next meeting Date.
  meetingDate = newMeetingDate(sheetDate, meetingNumber)

  //Meeting Date details
    var meetingDateYear= meetingDate.getFullYear()
    var meetingDateMonth = meetingDate.getMonth()
    var meetingDateDate = meetingDate.getDate()

  // Create New Agenda Sheet
  meetingDateFormated = Utilities.formatDate(meetingDate,TIMEZONE,DATE_FORMAT)
  meetingNumber = meetingNumber + 1


  var newAgendaSheet = ActiveSpreadsheet.insertSheet('#'+ meetingNumber +' | '+ meetingDateFormated,0,{template: AGENDA_TEMPLATE_SHEET})

  // Sets Date Value on the new Agenda.
  newAgendaSheet.getRange(SHEET_DATE_CELL).setValue(meetingDateFormated)

  // Sets Number of Meeting Value on the new Agenda.
  newAgendaSheet.getRange(MEETIING_NUMBER_CELL).setValue(meetingNumber)

  // Gets new meeting's Agenda URL.
  meetingAgendaURL = SpreadsheetApp.getActive().getUrl()


  // Creates the Meeting Notes Folder if it doesn't allready exists.
  var notesFolderID = createNewFolder(destinationFolderID, meetingTitle + " - Notes").getId()
  
  // Creates Meeting Note File.
  meetingNotesDoc = DriveApp.getFileById(NOTES_TEMPLATE_DOC_ID).makeCopy(meetingTitle +" Notes #"+ meetingNumber +" | "+ meetingDateFormated, DriveApp.getFolderById(notesFolderID))
  // Get Notes Doc URL.
  meetingNotesDocURL = meetingNotesDoc.getUrl()
  // Puts the Note URL on the new Agenda.
  linkCellContents('🔗 Meeting Notes link', meetingNotesDocURL, newAgendaSheet, MEETING_NOTES_LINK_CELL)
  // Replaces placeholders with meeting information.
  replacePlaceholdersInNotes(meetingTitle, meetingNumber, meetingDate, meetingDateFormated, meetingAgendaURL, meetingNotesDoc, Utilities.formatDate(meetingStartTime, TIMEZONE, "HH:mm"), Utilities.formatDate(meetingEndTime, TIMEZONE, "HH:mm"), EVENT_LOCATION)

  // Create Google Calendar Event Object.
  var meetingEventObj = calendraEvent( meetingTitle, EVENT_DESCRIPTION, EVENT_LOCATION, meetingDate, meetingStartTime, meetingEndTime, meetingNumber, meetingAgendaURL, meetingNotesDocURL, EVENT_GUESTS)
  // Create Google Calendar Event
  var newMeetingEvent = Calendar.Events.insert( meetingEventObj, CALENDAR_ID, {
   supportsAttachments: true,
   conferenceDataVersion: 1
   })
  var newMeetingEventID = newMeetingEvent.getId()
  
  // Gets the meeting Google Meet URL.
  meetingMeetURL = newMeetingEvent.hangoutLink
  // Puts the Google Meet URL on the new Agenda.
  linkCellContents('🔗 Google Meet link', meetingMeetURL, newAgendaSheet, MEETING_URL_CELL)
  // Puts the URL of the Note Folder.
  linkCellContents('🔗 Meeting Notes Folder 📂', DriveApp.getFolderById(notesFolderID).getUrl(), newAgendaSheet, MEETING_NOTES_FOLDER_LINK_CELL)
  //
  linkCellContents('🔗 Meeting Calendar 📆', CALENDAR_URL, newAgendaSheet, MEETING_CALENDAR_LINK_CELL)
}