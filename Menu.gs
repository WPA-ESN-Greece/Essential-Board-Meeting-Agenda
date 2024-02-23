function onOpen() 
{
  var ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu("🌌 ESN Menu")

  if (AGENDA_TEMPLATE_SHEET.getRange('C1').getValue() == 'Needs set-up')
  {
    menu.addItem("🔨 Set Up","initialSetup")
  }
  else
  {
    menu.addItem("📆 Create New Meeting Essentials","newMeetingEssentialsFromMenu")
  }

  menu.addItem("📚 View Documentation","showDocumentation")
  menu.addToUi()
}

function newMeetingEssentialsFromMenu()
{
  newMeetingEssentials("Menu")
}

//Documentation Link pop-up
function showDocumentation()
{
  let documentationMessage = HtmlService.createHtmlOutput(`<p style="font-family: 'Open Sans'">You can find the documentation <a href="${DOCUMENTATION_URL}"target="_blank">here</a>.</p>`).setWidth(400).setHeight(60)

  SpreadsheetApp.getUi().showModalDialog(documentationMessage,"📚 Documentation")
}