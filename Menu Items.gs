function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CloudFire')
  .addItem('Update Instructions and Promotion Planning Calendars for Teams', 'updateInstructionsAndPromotionPlanningCalendarsForTeams_v4')
  
  .addSeparator()
  
  .addSubMenu(
    SpreadsheetApp.getUi().createMenu('Contacting Team Leaders Before Deadline')
    .addItem("Invite Team Leaders to Contribute by the first Friday of October", "inviteTeamLeaderstoContributebythefirstFridayOfOctober")
    .addItem('Nudge Team Leaders Who Haven\'t Contributed Anything After a Week or Two', 'emailTeamLeadersWhoHaveNotDoneAnything')
    .addItem("Remind All Team Leaders to Contribute 'by this coming Friday'", "remindAllTeamLeadersToContributeByThisFriday")
    )
    
    .addSeparator()
    
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Contacting Team Leaders After Deadline')
      .addItem("Email Naughty and Nice Team Leaders", "emailNaughtyandNice")
    )
    
    .addToUi();
    }