function onOpen() 
{
  var ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu("🌌 ESN Menu")
  menu.addItem("📆 Create New Meeting Essentials","newMeetingEssentialsFromMenu").addToUi()
}

function newMeetingEssentialsFromMenu()
{
  newMeetingEssentials("Menu")
}