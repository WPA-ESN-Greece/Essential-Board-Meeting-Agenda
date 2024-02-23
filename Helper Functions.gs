// Gets the ID of a google doc file (Doc, spredsheet, presentation, form), folder or script from its URL.
function extractDocumentIdFromUrl(url) 
{
  var parts = url.split('/')

  if (parts[4] == "d")
  {
    var idIndex = parts.indexOf('d') + 1

    if (idIndex > 0 && idIndex < parts.length) 
    {
      return parts[idIndex]
    } 
    else 
    {
      // If the URL doesn't contain the expected parts
      Logger.log("Invalid URL")
      return "Invalid URL"
    }
  }

  if (parts[4] == "folders" || parts[4] == "projects" )
  {
    var idIndex = 5
    Logger.log(parts = url.split('/'))

    if (idIndex > 0 && idIndex < parts.length) 
    {
      return parts[idIndex]
    }
    else 
    {
      // If the URL doesn't contain the expected parts
      Logger.log("Invalid URL")
      return "Invalid URL";
    }
  }

  else
  {
    Logger.log("Unknown type of URL")
    return "Unknown type of URL"
  }
}


// Creates a Google Callendar Event object
function calendraEvent(_meetingName, _eventDestination, _eventLocation, _Date, _startTime, _endTime, _meetingNumber, _meetingAgendaUrl, _meetingNotesUrl, _guestEmail)
{
  // Google Calenda API Event Object: https://developers.google.com/calendar/api/v3/reference/events

  let eventStartTimeNdate = new Date(_Date.getFullYear(), _Date.getMonth(), _Date.getDate(), _startTime.getHours(), _startTime.getMinutes())
  let eventEndTimeNdate = new Date(_Date.getFullYear(), _Date.getMonth(), _Date.getDate(), _endTime.getHours(), _endTime.getMinutes())

  //Google Calendar Event Object and all of its properties. The commented ones are not in use. 
  let event = {
    //kind: "calendar#event",
    //"etag": etag,
    //"id": string,
    //"status": string,
    //"htmlLink": string,
    //"created": datetime,
    //"updated": datetime,
    summary: String(_meetingName) + " #" + String(_meetingNumber),
    description: String(_eventDestination),
    location: String(_eventLocation),
    colorId: 1,
    creator: {
      //"id": string,
      email: Session.getActiveUser().getEmail(),
      //"displayName": string,
      //"self": boolean
    },
    /*"organizer": {
      "id": string,
      "email": string,
      "displayName": string,
      "self": boolean
    },*/
    start: {
      dateTime: eventStartTimeNdate.toISOString()
      //date: String(_Date),
      //dateTime: String(_startTime),
      //timeZone: TIMEZONE
    },
    end: {
      dateTime: eventEndTimeNdate.toISOString()
      //date: String(_Date),
      //dateTime: String(_endTime),
      //timeZone: TIMEZONE
    },
    transparency: "opaque",
    visibility: "default",
    /*
    attendees: [
      {email: String(_guestEmail[0])},
      {email: String(_guestEmail[1])},
      {email: String(_guestEmail[2])},
      {email: String(_guestEmail[3])},
      {email: String(_guestEmail[4])},
    ].filter(attendees =>  attendees.email),
    */
    //"attendeesOmitted": boolean,
    /*"extendedProperties": {
      "private": {
        key: string
      },
      "shared": {
        key: string
      }
    },*/
    //hangoutLink: string,
    //conferenceData: {
      //conferenceDataVersion: 1,//
      //createRequest: {
        //requestId: "meet"+_meetingNumber,
        //conferenceSolutionKey: {
          //type: "hangoutsMeet"
        //},
        /*"status": {
          "statusCode": string
        }*/
      //},
      /*entryPoints: [
        {
          entryPointType: "video",
          "uri": string,
          "label": string,
          "pin": string,
          "accessCode": string,
          "meetingCode": string,
          "passcode": string,
          "password": string
        }
      ],*/
      /*conferenceSolution: {
        key: {
          type: "hangoutsMeet"
        },
        name: _meetingName + " #" + _meetingNumber,
        //"iconUri": string
      },*/
      //"conferenceId": string,
      //"signature": string,
      //"notes": string,
    //},
    //"anyoneCanAddSelf": boolean,
    //sendInvites: true,//
    guestsCanInviteOthers: true,
    guestsCanModify: true,
    guestsCanSeeOtherGuests: true,
    //"privateCopy": boolean,
    //"locked": boolean,
    reminders: {
      useDefault: false,
      overrides: [
        {method: 'popup', 'minutes': 24*60},
        {method: 'popup', 'minutes': 30},
        {method: 'popup', 'minutes': 15}
      ]
    },
    source: {
      url: String(_meetingAgendaUrl),
      title: "🚀 Meeting Agenda Automation Script."
    },
    /*workingLocationProperties: {
      type: "customLocation",
      //"homeOffice": (value),
      customLocation: {
        label: "📞 On an Online Meeting 💻"
      },
      "officeLocation": {
        "buildingId": string,
        "floorId": string,
        "floorSectionId": string,
        "deskId": string,
        "label": string
      }
    },*/
    //supportsAttachments:true,//
    attachments: [
      //Agenda
      {
        fileUrl: String(_meetingAgendaUrl),
        title: "📰 Agenda #" + String(_meetingNumber),
        mimeType: DriveApp.getFileById(extractDocumentIdFromUrl(_meetingAgendaUrl)).getMimeType(),
        //"iconLink": string,
        //"fileId": string
      },
      //Meeting
      {
        fileUrl: String(_meetingNotesUrl),
        title: "📝 Notes #" + String(_meetingNumber),
        mimeType: DriveApp.getFileById(extractDocumentIdFromUrl(_meetingNotesUrl)).getMimeType(),
        //iconLink: string,
        //"fileId": string
      }
    ],
    //eventType: "default"
  }
  
  // Adds every comma-seperated email address found in cell 'C10' (Meeting Guests) in the Template Sheet.
  if (_guestEmail.length > 0 && _guestEmail != "")
  {
    let attendeesArr = []

    for (var i = 0; i < _guestEmail.length; i++)
    {
      let emailAddressFormated = _guestEmail[i].trimStart()
      emailAddressFormated = _guestEmail[i].trim()
      
      attendeesArr.push({email:emailAddressFormated})
    }

    event.attendees = attendeesArr
  }
  
  // Gets the Meeting URL value foind in the Template Sheet. 
  let meetingURLinTemplate = AGENDA_TEMPLATE_SHEET.getRange(MEETING_URL_CELL).getValue()
  
  // Runs only if the the Meet URL Cell in the Template Sheet matches the text in 'MEET_URL_DEFAULT' found in the config.
  if (meetingURLinTemplate == MEET_URL_DEFAULT)
  {
    event.conferenceData = {
        //conferenceDataVersion: 1,
        createRequest: {
          requestId: "meet"+_meetingNumber,
          conferenceSolutionKey: {type: "hangoutsMeet"}
        }
      }
  }
  // Runs only if the the Meet URL Cell in the Template Sheet does NOT matche the text in 'MEET_URL_DEFAULT' found in the config AND is not empty. 
  else if (meetingURLinTemplate != "" && meetingURLinTemplate != MEET_URL_DEFAULT)
  {
    // Adds the user entered fixed meeting URL in the Events description. 
    event.description += `

    Meeting Link: ${meetingURLinTemplate}`
  }

  return event
}


// Create new folder function while it checks if it already exists.
function createNewFolder(parentFolderID, newFolderName)
{
  
  var folder = DriveApp.getFolderById(parentFolderID).getFolders()

  while(folder.hasNext()) 
  {
    var folderN = folder.next()

    if(folderN.getName() == newFolderName)
    {
      return folderN 
    }
  }
  var destinationFolder = DriveApp.getFolderById(parentFolderID).createFolder(newFolderName)
  
  return destinationFolder
}



//Day of the week function. Returns a numeric value that coresponds to a days of the week.

function dayOfTheWeek(string) 
{

  switch (string)
  {
    case "Monday":
      return 1

    case "Tuesday":
      return 2

    case "Wednesday":
      return 3

    case "Thursday":
      return 4      

    case "Friday":
      return 5    

    case "Saturday":
      return 6    

    case "Sunday":
      return 0 
  }
}


// A function that calculates the next day of the week. In this case, day = Tuesday.
function nextDay(date, day) {

  const result = new Date()//date.getTime()
  const offset = (((day + 6) - date.getDay()) % 7) + 1
  
  result.setDate(date.getDate() + offset)
  
  var dayr = result.getDate()
  var month = result.getMonth()
  var year = result.getFullYear()

  var resultFormated = new Date(year,month,dayr,null,null,null,null) //Utilities.formatDate(result,'GMT+2','d/M/YYYY')

  return resultFormated
}


// Calulates the next meeting date based on the chosen day of the week and returns the final meeting date.

function newMeetingDate(currentSheetDate = new Date("09/09/23"), _meetingNumber = 1) 
{
  
  
  var finalMeetingDate = Date()
  var currentDate = new Date()
  
  //Current Date details
     var currentDateYear = currentDate.getFullYear()
     var currentDateMonth = currentDate.getMonth()
     var currentDateDate = currentDate.getDate()

  currentDate = new Date( currentDateYear, currentDateMonth, currentDateDate, null, null, null, null)

  // 
  var DAY_OF_THE_WEEK = String(AGENDA_TEMPLATE_SHEET.getRange(DAY_OF_THE_WEEK_CELL).getValue())
  var weekDay = dayOfTheWeek(DAY_OF_THE_WEEK)

  //Calculates date of the next day of the week given. Example: If day of the week is Tuesday, it will calculate the next Tuesday date.
   var nextMeetingDate = new Date(nextDay(currentDate, weekDay))

  //NewMeetingDate details
    var nextMeetingDateDate = nextMeetingDate.getDate()
    var nextMeetingDateMonth = nextMeetingDate.getMonth()
    var nextMeetingDateYear = nextMeetingDate.getFullYear()

  //Control Point for new meeting date.
  if(currentSheetDate.valueOf() < currentDate.valueOf() && nextMeetingDate.valueOf() > currentDate.valueOf() && currentDate.getDay() == weekDay){
    finalMeetingDate = new Date(currentDateYear, currentDateMonth, currentDateDate)
    Logger.log("if")
    Logger.log(finalMeetingDate)
  }
  else if(currentSheetDate.valueOf() >= nextMeetingDate.valueOf() && _meetingNumber != 0){
   finalMeetingDate = new Date(nextMeetingDateYear, nextMeetingDateMonth, nextMeetingDateDate + 7 * _meetingNumber)
   Logger.log("else if")
   Logger.log(finalMeetingDate)
  }
  else{
    finalMeetingDate = new Date(nextMeetingDateYear, nextMeetingDateMonth, nextMeetingDateDate)
    Logger.log("else")
    Logger.log(finalMeetingDate)
  }
  return finalMeetingDate
}


function replacePlaceholdersInNotes(_meetingName, _meetingNumber, _meetingDate, _formatedDate, _agendaUrl, _meetingNotesDoc, _startTime, _endTime, _eventLocation)
{
 //Text Replacing for Notes
 var notesDoc = DocumentApp.openById(_meetingNotesDoc.getId())
 var body = DocumentApp.openById(notesDoc.getId()).getBody() 

  body.replaceText('{{Meeting Name}}', _meetingName)
  body.replaceText('{{Number}}', _meetingNumber)
  body.replaceText('{{Date}}', _formatedDate)
  body.findText('{{📰 Agenda Link}}').getElement().asText().setText('📰 Agenda').setLinkUrl( _agendaUrl).setFontSize(10).setForegroundColor('#000000').setBold(false).setUnderline(true)
  body.replaceText('{{Start Time}}', _startTime)
  body.replaceText('{{End Time}}', _endTime)
  body.replaceText('{{Location}}', _eventLocation)

  notesDoc.saveAndClose()
}

//A function that creates a stylised link on a cell on a sheet.

/**
 * Links the contents of a cell in a Google Sheets spreadsheet to a specified URL
 * with custom label and formatting.
 *
 * @param {string} label - The label text that will be displayed in the linked cell.
 * @param {string} url - The URL to which the cell content will be linked.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Google Sheets sheet containing the cell.
 * @param {string} cell - The cell address (A1 notation) where the linked content will be placed.
 * @returns {void}
 *
 * ```javascript
 * // Example usage:
 * var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 * var label = "Visit Google";
 * var url = "https://www.google.com";
 * var cell = "A1";
 * linkCellContents(label, url, sheet, cell);
 * ```
 */
function linkCellContents(label,url,sheet,cell) 
{
 var range = sheet.getRange(cell)
 
 var style = SpreadsheetApp.newTextStyle()
      .setItalic(false)
      .setBold(true)
      .setFontFamily("Roboto")
      .setFontSize(10)
      .setForegroundColor('#1155cc')
      .setUnderline(true)
      .build()
 
 var richValue = SpreadsheetApp.newRichTextValue()
 .setText(label)
 .setLinkUrl(url)
 .setTextStyle(style)
   
 range.setRichTextValue(richValue.build());
}


//Confirmation Alert
/**
 * Displays a custom alert dialog box in Google Apps Script.
 *
 * This function creates and displays an alert dialog with the specified title, message, and
 * customizable button options. It returns the user's response to the dialog.
 * @param {string} title - The title to display in the dialog box.
 * @param {string} message - The message to display in the dialog box.
 * @param {GoogleAppsScript.Base.Ui.ButtonSet} buttonsSet - The set of buttons to display in the dialog box.
 *   Possible values are:
 *   - `ui.ButtonSet.OK`: Display an OK button.
 *   - `ui.ButtonSet.OK_CANCEL`: Display OK and Cancel buttons.
 *   - `ui.ButtonSet.YES_NO`: Display Yes and No buttons.
 *   - `ui.ButtonSet.YES_NO_CANCEL`: Display Yes, No, and Cancel buttons.
 *
 * @returns {GoogleAppsScript.Base.Ui.Button} The button that was clicked in the dialog box.
 */
function showAlert(title, message, buttonsOptions)
{
  var ui = SpreadsheetApp.getUi()
  var response  = ui.alert(String(title), String(message), buttonsOptions)
  return response
}


function isDateValid(datetimeString)
{
  var datetimeString = "25/02/2024, 15:30-17:00"

  // Date format dd/MM/yyyy
  const expectedDateRegEx = /^\d{2}\/\d{2}\/\d{4}, \d{2}:\d{2}-\d{2}:\d{2}$/
  const Demo_datetimeString = "15/09/2023, 15:30-17:00"
  
  if (datetimeString.length < Demo_datetimeString.length) {return {status: false}}
  
  let date = String(datetimeString).split(", ")[0]
  let timeStart = String(datetimeString).split(", ")[1].split("-")[0]
  let timeEnd = String(datetimeString).split(", ")[1].split("-")[1]

  // Test the input string against the pattern
  let match = datetimeString.match(expectedDateRegEx)

  Logger.log(`-------- date regex match: ${match} `)

  if (!match) 
  {
    Logger.log(`The string '${datetimeString}' format doesn't match "dd/MM/yyyy, HH:mm-HH:mm"`)
    return outputObj = {status: false}; // The string format doesn't match "dd/MM/yyyy"
  }
  
  Logger.log("date " + date)

  var dateInQuestion = new Date(date.split("/")[2], date.split("/")[1] - 1, date.split("/")[0], null,null,null,null)
  
  Logger.log("dateInQuestion " + dateInQuestion)
  
  var day = dateInQuestion.getDate()
  var month = dateInQuestion.getMonth() +1
  var year = dateInQuestion.getFullYear()

  Logger.log(`Day: ${day}, Month: ${month}, Year: ${year}`)

  // (((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0) == true && (month == 2 && day > 29)) || ([4, 6, 9, 11].includes(month) == true && day > 30)) 
    // ((true || false) && false) || ([4, 6, 9, 11].includes(month) && day > 30)) 
  Logger.log("========")
  Logger.log(((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)) && ((month == 2 && day > 29) || ([4, 6, 9, 11].includes(month) && day > 30)))
  Logger.log("========")

  // Check if the extracted day and month are within valid ranges
  if (day < 1 || day > 31 || month < 1 || month > 12) {
    Logger.log("Day or month is out of range.")
    return outputObj = {status: false} // Day or month is out of range
  }
  else if (((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0) == false) && (month == 2 && day > 28)) 
  {
    Logger.log("Febuary has more than 28 days in a non-leap year.")
    return outputObj = {status: false}
  }
  else if (((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)) && ((month == 2 && day > 29) || ([4, 6, 9, 11].includes(month) && day > 30))) 
  {
    Logger.log("Febuary has more than 29 days in a leap year OR a 20 day month has more than 30 days.")
    return outputObj = {status: false}
  }
  else 
  { 
    Logger.log("Valid Date.")
    outputObj = {
      status: true, 
      startDate: new Date(year, month -1, day, timeStart.split(":")[0], timeStart.split(":")[1], null, null),
      endDate: new Date(year, month -1, day, timeEnd.split(":")[0], timeEnd.split(":")[1], null, null)
      }
    Logger.log(outputObj)
    return outputObj
  }
}



function getWeekDayForTrigger()
{
  let weekDayToTrigger = 0
  let weekDayNumber = dayOfTheWeek(DAY_OF_THE_WEEK)

  if (weekDayNumber >= 0 && weekDayNumber < 6)
  {
    weekDayToTrigger = weekDayNumber + 1
  }
  else {weekDayToTrigger = 0}
  
  switch (weekDayToTrigger)
  {
    case 1:
      return ScriptApp.WeekDay.MONDAY
    case 2:
      return ScriptApp.WeekDay.TUESDAY
    case 3:
      return ScriptApp.WeekDay.WEDNESDAY
    case 4:
      return ScriptApp.WeekDay.THURSDAY      
    case 5:
      return ScriptApp.WeekDay.FRIDAY    
    case 6:
      return ScriptApp.WeekDay.SATURDAY      
    case 0:
      return ScriptApp.WeekDay.SUNDAY    
  }

}
