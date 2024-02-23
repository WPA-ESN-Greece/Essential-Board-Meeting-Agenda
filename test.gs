function myFunction() {
  
  
  
  
  
  
  

















  var guests = ss.getRange(MEETING_GUESTS_CELL).getValue().split(",")

  EVENT_GUESTS.forEach((emailAddress, index) => {
    EVENT_GUESTS[index] = emailAddress.trimStart();
    EVENT_GUESTS[index] = emailAddress.trim();})

  Logger.log(guests)
}
