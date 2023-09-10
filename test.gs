/*
function myFunction() {
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


  var meetingStartTime123 = new Date(null,null,null,START_TIME_HOURS,START_TIME_MINUTES)


  Logger.log(meetingStartTime123.getMinutes())
  Logger.log(Utilities.formatString('%00s',meetingStartTime123.getMinutes()))
  Logger.log(Utilities.formatDate(meetingStartTime123, TIMEZONE, "HH:mm"))
}
*/