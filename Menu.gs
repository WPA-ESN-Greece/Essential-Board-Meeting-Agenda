function onOpen() 
{
  var ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu("🌌 ESN Menu")

  if (AGENDA_TEMPLATE_SHEET.getRange('C1').getValue() == 'Needs set-up')
  {
    menu.addItem("🔨 Set Up","initialSetup").addToUi()
  }
  else
  {
    menu.addItem("📆 Create New Meeting Essentials","newMeetingEssentialsFromMenu").addToUi()
  }

}

function newMeetingEssentialsFromMenu()
{
  newMeetingEssentials("Menu")
}