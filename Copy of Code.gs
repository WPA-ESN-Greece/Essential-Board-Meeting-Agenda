/*
function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("ESN Menu")
  menu.addItem("Create New Meeting","newMeeting")
  menu.addToUi()
}

const ActiveSheet = SpreadsheetApp.getActiveSpreadsheet()

const AGENDA_TEMPLATE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('#No | Date')
const NOTES_TEMPLATE_DOC = AGENDA_TEMPLATE_SHEET.getRange('L9').getValue()
const NOTES_TEMPLATE_DOC_ID = DriveApp.getFileById(NOTES_TEMPLATE_DOC)

const MEETIING_NAME = ActiveSheet.getName().split(" |",1) //AGENDA_TEMPLATE_SHEET.getRange('L12').getValue()

const START_TIME = AGENDA_TEMPLATE_SHEET.getRange('L19').getValue()
  const START_TIME_HOURS = START_TIME.split(":",1)
  const START_TIME_MINUTES = START_TIME.split(":",2).slice(1,2)

const END_TIME = AGENDA_TEMPLATE_SHEET.getRange('L21').getValue()
  const END_TIME_HOURS = END_TIME.split(":",1)
  const END_TIME_MINUTES = END_TIME.split(":",2).slice(1,2)

const DATE_FORMAT = "dd/MM/yy"
const DAY_OF_THE_WEEK = String(AGENDA_TEMPLATE_SHEET.getRange('L15').getValue())
const CALENDAR_ID = String(AGENDA_TEMPLATE_SHEET.getRange('L17').getValue())

const EVENT_DESCRIPTION = AGENDA_TEMPLATE_SHEET.getRange('L6').getValue()
const EVENT_LOCATION = "Google Meet"
const EVENT_GUESTS = [
  AGENDA_TEMPLATE_SHEET.getRange('L24').getValue(),
  AGENDA_TEMPLATE_SHEET.getRange('L25').getValue(),
  AGENDA_TEMPLATE_SHEET.getRange('L26').getValue(),
  AGENDA_TEMPLATE_SHEET.getRange('L27').getValue(),
  AGENDA_TEMPLATE_SHEET.getRange('L28').getValue()].filter(n => n)          
 
  const SLACK_CHANNEL_EMAIL = "wpa+slack@esngreece.gr"

function newMeeting(){
  var timeZone = Session.getTimeZone()

  var CurrentAgenda = ActiveSheet.getActiveSheet()

  var SHEET_DATE = new Date(CurrentAgenda.getRange('L4').getValue())
  
  var destinationFolderID = DriveApp.getFileById(ActiveSheet.getId()).getParents().next().getId() //Agenda Parent folder

  //Meeting Object
  let Meeting = {
    date: SHEET_DATE,
    formatedDate: "", 
    title: MEETIING_NAME,
    notes:"",
    notesUrl:"",
    agendaUrl:"",
    meetUrl:"",
    number: CurrentAgenda.getRange('L3').getValue(),
    startTime: new Date(null,null,null,START_TIME_HOURS,START_TIME_MINUTES),
    endTime: new Date(null,null,null,END_TIME_HOURS,END_TIME_MINUTES),
    standardDay: DAY_OF_THE_WEEK 
  }

  //Current Date
   var currDateRaw = new Date()
   var currDate = removeTime(currDateRaw)
    
    //Current Date details
     var currDateYear = currDate.getFullYear()
     var currDateMonth = currDate.getMonth()
     var currDateDate = currDate.getDate()

  var weekDay = dayOfTheWeek(DAY_OF_THE_WEEK)

  //Calculates date of the next Tuesday.
   var nextMeetDay = nextDay(currDate, weekDay)
   //Logger.log(nextMeetDay + " nextMeetDate")
   var newMeetingDate = new Date(nextMeetDay)

   //NewMeetingDate details
    var newMeetingDateDate = newMeetingDate.getDate()
    var newMeetingDateMonth = newMeetingDate.getMonth()
    var newMeetingDateYear = newMeetingDate.getFullYear()
  
 //Control Point for new meeting date.
  if(Meeting.date.valueOf() < currDate.valueOf() && newMeetingDate.valueOf() > currDate.valueOf() && currDate.getDay() == weekDay){
    Meeting.date = new Date(currDateYear, currDateMonth, currDateDate)
    Logger.log("if")
    Logger.log(Meeting.date)
  }
  else if(Meeting.date.valueOf() >= newMeetingDate.valueOf() && Meeting.number != 0){
   Meeting.date = new Date(newMeetingDateYear, newMeetingDateMonth, newMeetingDateDate + 7 * Meeting.number)
   Logger.log("else if")
   Logger.log(Meeting.date)
  }
  else{
    Meeting.date = new Date(newMeetingDateYear, newMeetingDateMonth, newMeetingDateDate)
    Logger.log("else")
    Logger.log(Meeting.date)
  }


//Meeting Date details
  var meetDateYear= Meeting.date.getFullYear()
  var meetDateMonth = Meeting.date.getMonth()
  var meetDateDate = Meeting.date.getDate()


//Creating new sheet
 Meeting.formatedDate = Utilities.formatDate(Meeting.date,timeZone,DATE_FORMAT)
 Meeting.number++

 var NewAgenda = ActiveSheet.insertSheet('#'+ Meeting.number +' | '+ Meeting.formatedDate,0,{template: AGENDA_TEMPLATE_SHEET})

 NewAgenda.getRange('L4').setValue(Meeting.formatedDate)
 NewAgenda.getRange('L3').setValue(Meeting.number)
 
 Meeting.agendaUrl = SpreadsheetApp.getActive().getUrl()

 
//Getting new notes document
var notesFolderID = createNewFolder(destinationFolderID , Meeting.title+" - Notes").getId()


Meeting.notes = DriveApp.getFileById(NOTES_TEMPLATE_DOC).makeCopy(Meeting.title +" Notes #"+ Meeting.number +" | "+ Meeting.formatedDate,DriveApp.getFolderById(notesFolderID))

Meeting.notesUrl = Meeting.notes.getUrl()
linkCellContents('  📝 Meeting Notes',Meeting.notesUrl,NewAgenda,'A2:I2')

//Text Replacing for Notes
 var notesDoc = DocumentApp.openById(Meeting.notes.getId())
 var body = DocumentApp.openById(notesDoc.getId()).getBody()

  body.replaceText('{{Meeting Name}}',Meeting.title)
  body.replaceText('{{Number}}',Meeting.number)
  body.replaceText('{{Date}}',Meeting.formatedDate)
  body.findText('{{📰 Agenda}}').getElement().asText().setText('📰 Agenda').setLinkUrl(Meeting.agendaUrl).setFontSize(11).setForegroundColor('#000000').setBold(false).setUnderline(true)

  notesDoc.saveAndClose()


//Calendar Event 
 var eventTitle = Meeting.title +" #"+ Meeting.number
 var eventStartTimeNdate = new Date(meetDateYear, meetDateMonth, meetDateDate, Meeting.startTime.getHours(),Meeting.startTime.getMinutes())
 var eventEndTimeNdate = new Date(meetDateYear, meetDateMonth, meetDateDate, Meeting.endTime.getHours(),Meeting.endTime.getMinutes())

  var event = {
    summary: eventTitle,
    location: EVENT_LOCATION,
    description: EVENT_DESCRIPTION,
    start: {
      dateTime: eventStartTimeNdate.toISOString()
    },
    end: {
      dateTime: eventEndTimeNdate.toISOString()
    },
    attendees: [
      {email: EVENT_GUESTS[0]},
      {email: EVENT_GUESTS[1]},
      {email: EVENT_GUESTS[2]},
      {email: EVENT_GUESTS[3]},
      {email: EVENT_GUESTS[4]},
    ].filter(attendees =>  attendees.email),

    colorId: 1,
    reminders: {
     useDefault: false,
     overrides: [
      {method: 'email', 'minutes': 3*24*60},
      {method: 'popup', 'minutes': 30},
      {method: 'popup', 'minutes': 15}
    ]},
    supportsAttachments:true,
    attachments:[
    {'fileUrl': Meeting.agendaUrl,
    'title':'📰 Agenda #' + Meeting.number },
    {'fileUrl': Meeting.notesUrl,
    'title': '📝 Notes #' + Meeting.number}
    ],
    sendInvites: true,
    guestsCanInviteOthers:true,
    guestsCanModify: true,
    guestsCanSeeOtherGuests: true,
    

    conferenceData: { 
      //conferenceDataVersion: 1,
      createRequest: {
      requestId: "meet"+Meeting.number,
      conferenceSolutionKey: {
        type: "hangoutsMeet",
      },
      
     },
     
    }
  }

 var newEvent = Calendar.Events.insert(event, CALENDAR_ID,{
   supportsAttachments: true,
   conferenceDataVersion: 1
   })
 


 var newEventId = newEvent.getId() 

 //Getting the Google Meet Url of the event. 
 Meeting.meetUrl = newEvent.hangoutLink
 NewAgenda.getRange('A1').setRichTextValue(SpreadsheetApp.newRichTextValue().setText("📞").setLinkUrl(Meeting.meetUrl).build())

  var message = 
    `<h2>${Meeting.title} #${Meeting.number}</h2>`+
    '<p>Next meeting has been scheduled.</p>'+
    `<h3>for ${meetDateDate}/${meetDateMonth}/${meetDateYear} 📆</h3>`+
    `<p>(It's a ${DAY_OF_THE_WEEK})</p>`+
    `<p>Here is the <a href="${Meeting.agendaUrl} 📰">Agenda Link</a> so you can add topics.</p>`;
                                             

  MailApp.sendEmail(
    {
            to: SLACK_CHANNEL_EMAIL,
            //cc: ,
            subject: `New ${Meeting.title} is Ready!`,
            htmlBody: message,
    }
  )


 Logger.log("The new meeting link: " + Meeting.meetUrl)

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

//A function that creates a stylised link on a cell on a sheet.
function linkCellContents(label,url,sheet,cell) {
 var range = sheet.getRange(cell)
 var style = SpreadsheetApp.newTextStyle()
      .setItalic(false)
      .setFontSize(10)
      .setForegroundColor('#ffffff')
      .setUnderline(false)
      .build()
 var richValue = SpreadsheetApp.newRichTextValue()
 .setText(label)
 .setLinkUrl(url)
 .setTextStyle(style)
   
 range.setRichTextValue(richValue.build());
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

//A function that removes the time from a given date.
 function removeTime(date){

  var dateYear = date.getFullYear()
  var dateMonth = date.getMonth()
  var dateDate = date.getDate()

  var result = new Date(dateYear, dateMonth, dateDate,0,0,0,0)

  return result
}

//A function that calculates the next day of the week. In this case, day = Tuesday.
function nextDay(date, day) {

  const result = new Date()//date.getTime()
  const offset = (((day + 6) - date.getDay()) % 7) + 1
  
  result.setDate(date.getDate() + offset)
  
  var dayr = result.getDate()
  var month = result.getMonth()
  var year = result.getFullYear()

  var resultFormated = new Date(year,month,dayr,0,0,0,0) //Utilities.formatDate(result,'GMT+2','d/M/YYYY')

  return resultFormated
}
*/