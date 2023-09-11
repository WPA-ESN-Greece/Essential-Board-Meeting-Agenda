
function myFunction() 
{
  var ui = SpreadsheetApp.getUi()
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
  //var alertButtonOption = showAlert("A", "fff", ui.ButtonSet.YES_NO_CANCEL)
  
  var meetingDateToConfirm = new Date(meetingDateYear, meetingDateMonth, meetingDateDate, START_TIME_HOURS, START_TIME_MINUTES,null,null)
  var alertResponse = showAlert(
    "📆 About your new meeting",
    `You are about to create the essentials ✨ for a meeting on ${Utilities.formatDate(meetingDateToConfirm, TIMEZONE, "EEE dd/MM/yy 'at' HH:mm aaa z")} till ${END_TIME}. Do you wish to continue? To input a custom Date and time, click the "No" button.`, 
    ui.ButtonSet.YES_NO_CANCEL)
    Logger.log("--- alertResponse "+alertResponse)
    Logger.log(ui.Button.NO)
  
  if (alertResponse === ui.Button.YES)
  {
    Logger.log("Yes!")
  }
  else if (alertResponse === ui.Button.CANCEL || alertResponse === ui.Button.CLOSE)
  {
    Logger.log("Cancel/ Close")
    return;
  }
  else if (alertResponse === ui.Button.NO)
  {
    Logger.log("No")
    var input = ui.prompt("When would you like to scedule your meeting? Your answer HAS TO be in the following date & time format: dd/MM/yyyy, HH:mm-HH:mm. That's date, starting time and end time.")
    var promptResponseText = input.getResponseText()
    var validDateOutputObj = isDateValid(promptResponseText)
    Logger.log("Is date valid: " + validDateOutputObj.status)
    Logger.log(validDateOutputObj.endDate)
    while (validDateOutputObj.status == false)
    {
      input = ui.prompt("Wrong format 😢. Please try again. When would you like to scedule your meeting? Your answer HAS TO be in the following date & time format: dd/MM/yyyy,HH:mm. Example: 20/09/2023, 13:00-15:30")
      promptResponseText = input.getResponseText()
      validDateOutputObj = isDateValid(promptResponseText)

      if (input.getSelectedButton() === ui.Button.CLOSE) {return}
    }
  }
  else {return}
  Logger.log("Creating  event...")
  
  //Logger.log("=== alertResponse " + alertResponse)
  /*if (alertButtonOption === ui.ButtonSet.YES) { Logger.log("Yes!") }
  else if (alertButtonOption === ui.ButtonSet.CANCEL) { Logger.log("Cancel"); return}
  else 
  {
    Logger.log("Need an input")
  }*/
  //Logger.log("Making an event...")
}
