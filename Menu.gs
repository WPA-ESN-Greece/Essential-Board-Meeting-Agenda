function onOpen() 
{
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("🌌 ESN Menu")
  menu.addItem("📆 Create New Meeting Essentials","newMeetingEssentials").addToUi()
}

