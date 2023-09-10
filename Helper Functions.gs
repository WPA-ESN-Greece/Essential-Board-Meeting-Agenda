// Gets the ID of a google doc file (Doc, spredsheet, presentation, form), folder or script from its URL.

function extractDocumentIdFromUrl(url) 
{
  var parts = url.split('/')

  if (parts[4] == "d")
  {
    var idIndex = parts.indexOf('d') + 1
    //Logger.log(parts = url.split('/'))

    if (idIndex > 0 && idIndex < parts.length) 
    {
      //Logger.log(parts[idIndex])
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
      //Logger.log(parts[idIndex])
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

  var eventStartTimeNdate = new Date(_Date.getFullYear(), _Date.getMonth(), _Date.getDate(), _startTime.getHours(), _startTime.getMinutes())
  var eventEndTimeNdate = new Date(_Date.getFullYear(), _Date.getMonth(), _Date.getDate(), _endTime.getHours(), _endTime.getMinutes())

  //Google Calendar Event Object
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
    attendees: [
      {
        email: String(_guestEmail[0]),
      }
    ],
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
    conferenceData: {
      //conferenceDataVersion: 1,//
      createRequest: {
        requestId: "meet"+_meetingNumber,
        conferenceSolutionKey: {
          type: "hangoutsMeet"
        },
        /*"status": {
          "statusCode": string
        }*/
      },
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
    },
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
  
  /*var newMeetingEvent = Calendar.Events.insert(event, CALENDAR_ID)
  var newMeetingEventID = newMeetingEvent.getId()
  
 var eventDetails = {
  id: newMeetingEvent.getId(),
  googleMeetURL: newMeetingEvent.hangoutLink
 }*/

  return event
}


//Create new folder function while it checks if it already exists.
function createNewFolder(parentFolderID, newFolderName){
  
  //var parentFolderID = parentFolder.getId()
  var folder = DriveApp.getFolderById(parentFolderID).getFolders()

  while(folder.hasNext()) {
    var folderN = folder.next()
    if(folderN.getName() == newFolderName){
      return folderN }
  }
  var destinationFolder = DriveApp.getFolderById(parentFolderID).createFolder(newFolderName)
  return destinationFolder
}



//Day of the week function. Returns a numeric value that coresponds to a days of the week.

function dayOfTheWeek(string) {
  if(string === 'Monday') return 1
  if(string === 'Tuesday') return 2
  if(string === 'Wednesday') return 3
  if(string === 'Thursday') return 4
  if(string === 'Friday') return 5
  if(string === 'Saturday') return 6
  if(string === 'Sunday') return 0
}


//A function that calculates the next day of the week. In this case, day = Tuesday.
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