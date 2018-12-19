function onOpen() {

  var ui = SpreadsheetApp.getUi();
  
  ui
    .createMenu('CloudFire')
    .addItem('Update Instructions and Promotion Planning Calendars for Teams', 'updateInstructionsAndPromotionPlanningCalendarsForTeams_v4')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Contacting Team Leaders Before Deadline')
        .addItem("Invite Team Leaders to Contribute by the first Friday of October", "inviteTeamLeaderstoContributebythefirstFridayOfOctober")
        .addItem("Nudge Team Leaders Who Haven\'t Contributed Anything After a Week or Two", 'emailTeamLeadersWhoHaveNotDoneAnything')
        .addItem("Remind All Team Leaders to Contribute 'by this coming Friday'", "remindAllTeamLeadersToContributeByThisFriday")
      )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Contacting Team Leaders After Deadline')
      .addItem("Email Naughty and Nice Team Leaders", "emailNaughtyandNice")
    )
    .addToUi();
}

// Menu items
function updateInstructionsAndPromotionPlanningCalendarsForTeams_v4() {PCT.updateInstructionsAndPromotionPlanningCalendarsForTeams_v4()}
function inviteTeamLeaderstoContributebythefirstFridayOfOctober()     {PCT.inviteTeamLeaderstoContributebythefirstFridayOfOctober()}
function emailTeamLeadersWhoHaveNotDoneAnything()                     {PCT.emailTeamLeadersWhoHaveNotDoneAnything()}
function remindAllTeamLeadersToContributeByThisFriday()               {PCT.remindAllTeamLeadersToContributeByThisFriday()}
function emailNaughtyandNice()                                        {PCT.emailNaughtyandNice()}

// Simple Triggers
function onEdit() {PCT.onEdit()}