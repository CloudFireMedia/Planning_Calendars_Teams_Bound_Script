/*
INTENT:
Updates the "Promotion Planning Calendars for Teams.gsheet" with instructions
and team information from the " Promotion Planning Calendar for Teams TEMPLATE.gsheet".
It also changes the date validation for Col B to "next year" and protects the ranges
of the current year"s event data on each team sheet.

-------------------------------------------------------
Redevelopment Notes from Chad:

1. 
//Notify Chad that the component is updated and aready for Team Leaders" contributions.
var PromotionPlanningCalendarsforTeams_url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
var subject = "Notice: Promotion Planning Calendars for Teams have been updated";
var body = Utilities.formatString("\
Communications Director: <br><br>\
The "Instructions" and "Team" tabs in the <a href="%s">Promotion Planning Calendars for Teams</a> component have been automatically updated for the current calendar year.<br><br>\
Team Leaders will be invited to comment on their Planning Calendars beginning tomorrow.<br><br>\
The deadline for contributions is the first Friday of September.\
",
PromotionPlanningCalendarsforTeams_url
);

MailApp.sendEmail({
name: "communications@ccnash.org",
to: "chbarlow@gmail.com",
subject: subject,
htmlBody: body
});

2. 
This script creates a 'team sheet' called 'N/A', which is captured from events on the Promo Deadlines Cal sponsored by 'N/A'.
Rather than this, we need the script to capture all events on the Promo Deadlines Cal that are not explicilty assigned to any 
team found on the Staff Data sheet and dump them into one sheet called 'Unsponsored Events'.

-------------------------------------------------------
*/
//This function is run once annually. Update the Instructions sheet and all Team sheets in reference to the current calendar year
function updateInstructionsAndPromotionPlanningCalendarsForTeams_v4() {
  var promotionPlanningCalendarsForTeamsTEMPLATE = SpreadsheetApp.openById(PROMOTION_PLANNING_CALENDARS_FOR_TEAMS_TEMPLATE_ID_);
  var promotionPlanningCalendarForTeams = SpreadsheetApp.openById(PROMOTION_PLANNING_CALENDAR_FOR_TEAMS_ID_);
  var staffSheet = SpreadsheetApp.openById(STAFF_DATA_ID_);
  //verify manual execution is intentional and not run too early
  var title = "Warning";
  var prompt = Utilities.formatString(
    "Executing this script will result in the deletion of all present team sheets.\n\n This will mean the loss of any data that may have been contributed manually to these sheets.\n\n Do you wish to proceed?"
  );
  if (SpreadsheetApp.getUi()
    .alert(title, prompt, Browser.Buttons.YES_NO) != "YES") return;
  //change date validation on Template sheet
  var currentYear = new Date()
    .getYear();
  var nextYear = new Date()
    .getYear() + 1;
  // Change existing data validation rules in row 1 and replicate down rows
  var oldDates = [new Date(currentYear, 0, 1), new Date(currentYear, 11, 31)];
  var newDates = [new Date(nextYear, 0, 1), new Date(nextYear, 11, 31)];
  var rangeDates = promotionPlanningCalendarsForTeamsTEMPLATE.getRange(
    "H4:H100");
  var rules = rangeDates.getDataValidations();
  for (var i = 0; i < rules.length; i++) {
    for (var j = 0; j < rules[i].length; j++) {
      var rule = rules[i][j];
      if (rule != null) {
        var criteria = rule.getCriteriaType();
        var args = rule.getCriteriaValues();
        if (criteria == SpreadsheetApp.DataValidationCriteria.DATE_BETWEEN &&
          args[0].getTime() == oldDates[0].getTime() &&
          args[1].getTime() == oldDates[1].getTime()) {
          // Create a builder from the existing rule, then change the dates.
          rules[i][j] = rule.copy()
            .withCriteria(criteria, newDates)
            .build();
        }
      }
    }
  }
  rangeDates.setDataValidations(rules);
  //get due date for Team Member contributions
  var weeks_array = getWeeksInMonth_();
  var teamLeadersContributionDueDate = getFirstWeek_();
  //update Planning Calendars for Teams "Instructions" sheet
  var instructionsTemplate_F4 = promotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName(
      "Instructions")
    .getRange("F4");
  var instructionsTemplate_D7 = promotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName(
      "Instructions")
    .getRange("D7");
  var instructionsTemplate_D8 = promotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName(
      "Instructions")
    .getRange("D8");
  var instructionsTemplate_D9 = promotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName(
      "Instructions")
    .getRange("D9");
  var instructions_F4 = promotionPlanningCalendarForTeams.getSheetByName(
      "Instructions")
    .getRange("F4");
  var instructions_D7 = promotionPlanningCalendarForTeams.getSheetByName(
      "Instructions")
    .getRange("D7");
  var instructions_D8 = promotionPlanningCalendarForTeams.getSheetByName(
      "Instructions")
    .getRange("D8");
  var instructions_D9 = promotionPlanningCalendarForTeams.getSheetByName(
      "Instructions")
    .getRange("D9");
  var instructionDetails_F4 = instructionsTemplate_F4.getValue()
    .toString();
  instructionDetails_F4 = instructionDetails_F4.replace(
    /TEAM_LEADERS_CONTRIBUTION_DUE_DATE/g, teamLeadersContributionDueDate);
  var instructionDetails_D7 = instructionsTemplate_D7.getValue()
    .toString();
  instructionDetails_D7 = instructionDetails_D7.replace(/UPCOMING_YEAR/g,
    currentYear + 1);
  var instructionDetails_D8 = instructionsTemplate_D8.getValue()
    .toString();
  instructionDetails_D8 = instructionDetails_D8.replace(/UPCOMING_YEAR/g,
    currentYear + 1);
  var instructionDetails_D9 = instructionsTemplate_D9.getValue()
    .toString();
  instructionDetails_D9 = instructionDetails_D9.replace(/UPCOMING_YEAR/g,
    currentYear + 1);
  instructions_F4.setValue(instructionDetails_F4);
  instructions_D7.setValue(instructionDetails_D7);
  instructions_D8.setValue(instructionDetails_D8);
  instructions_D9.setValue(instructionDetails_D9);
  // Get a list of team names 
  var TEAM_COLUMN_INDEX = 12;
  var EMPLOYEE_FIRST_RECORD = 3;
  var teamArray = [];
  var staff = staffSheet.getSheetByName("Staff");
  for (var i = EMPLOYEE_FIRST_RECORD; i < staff.getLastRow(); i++) {
    var teamName = staff.getRange(i, TEAM_COLUMN_INDEX)
      .getValue();
    if (teamName == '') teamName = 'N/A';
    if (teamArray.indexOf(teamName) == -1) {
      teamArray.push(teamName);
    }
  }
  var teams = teamArray.sort();
  //remove the existing team sheets
  var currentTeamSheets = promotionPlanningCalendarForTeams.getSheets();
  for (var i = 1; i < currentTeamSheets.length; i++) {
    var currentSheetName = currentTeamSheets[i].getName();
    if (currentSheetName.toLowerCase()
      .indexOf("team")) promotionPlanningCalendarForTeams.deleteSheet(
      currentTeamSheets[i]);
  }
  //enumerate the team names and add new sheets
  for (var newTeamSheetIndex = 0; newTeamSheetIndex < teams.length; newTeamSheetIndex++) {
    var teamName = teams[newTeamSheetIndex];
    var teamTemplate = promotionPlanningCalendarsForTeamsTEMPLATE.getSheetByName(
      "TEAM");
    var newTeamSheet = teamTemplate.copyTo(promotionPlanningCalendarForTeams);
    // copy formulas and replace TEAM and UPCOMING_YEAR in the formula
    var formulaRanges = teamTemplate.getRange(4, 1, teamTemplate.getLastColumn(),
      teamTemplate.getLastColumn());
    var formulas = formulaRanges.getFormulas();
    for (var i in formulas) {
      for (var j in formulas[i]) {
        var templateFormula = formulas[i][j];
        formulas[i][j] = templateFormula.replace(/TEAM/g, teamName);
        Logger.log(formulas[i][j]);
      }
    }
    for (var k in formulas) {
      for (var l in formulas[k]) {
        var templateFormula2 = formulas[k][l];
        formulas[k][l] = templateFormula2.replace(/UPCOMING_YEAR/g, currentYear +
          1);
      }
    }
    //set new values to variables in new sheets
    var newTeamRanges = newTeamSheet.getRange(4, 1, teamTemplate.getLastColumn(),
      teamTemplate.getLastColumn());
    newTeamRanges.setFormulas(formulas);
    newTeamSheet.setName(teamName);
    var newTeamSheetTitle = newTeamSheet.getRange("A1");
    newTeamSheetTitle.setValue(newTeamSheetTitle.getValue()
      .toString()
      .replace(/TEAM/g, teamName));
    newTeamSheetTitle.setValue(newTeamSheetTitle.getValue()
      .toString()
      .replace(/UPCOMING_YEAR/g, currentYear + 1));
    var newTeamSheetCurrentYear = newTeamSheet.getRange("A2");
    newTeamSheetCurrentYear.setValue(newTeamSheetCurrentYear.getValue()
      .toString()
      .replace(/CURRENT_YEAR/g, currentYear));
    var newTeamSheetUpcomingYear = newTeamSheet.getRange("C2");
    newTeamSheetUpcomingYear.setValue(newTeamSheetUpcomingYear.getValue()
      .toString()
      .replace(/UPCOMING_YEAR/g, currentYear + 1));
    var newTeamSheetCurrentYearPromoTier = newTeamSheet.getRange("A3");
    newTeamSheetCurrentYearPromoTier.setValue(newTeamSheetCurrentYearPromoTier.getValue()
      .toString()
      .replace(/CURRENT_YEAR/g, currentYear));
    var newTeamSheetCurrentYearStartDate = newTeamSheet.getRange("B3");
    newTeamSheetCurrentYearStartDate.setValue(newTeamSheetCurrentYearStartDate.getValue()
      .toString()
      .replace(/CURRENT_YEAR/g, currentYear));
    var newTeamSheetWillEventRecur = newTeamSheet.getRange("C3");
    newTeamSheetWillEventRecur.setValue(newTeamSheetWillEventRecur.getValue()
      .toString()
      .replace(/UPCOMING_YEAR/g, currentYear + 1));
    var newTeamProposedPromoTierForUpcomingYear = newTeamSheet.getRange("H3");
    newTeamProposedPromoTierForUpcomingYear.setValue(
      newTeamProposedPromoTierForUpcomingYear.getValue()
      .toString()
      .replace(/UPCOMING_YEAR/g, currentYear + 1));
    var newTeamProposedStartDateForUpcomingYear = newTeamSheet.getRange("I3");
    newTeamProposedStartDateForUpcomingYear.setValue(
      newTeamProposedStartDateForUpcomingYear.getValue()
      .toString()
      .replace(/UPCOMING_YEAR/g, currentYear + 1));
    
    var teamLeadersContributionDueDateinString = teamLeadersContributionDueDate + "";
    var teamLeaderContributionDueDateMinusOrdinand = teamLeadersContributionDueDateinString.slice(0, -2);
    
    var teamLeaderContributionDueDateMinusOrdinandPlusYearAsReferenceForConditionalFormatting = newTeamSheet.getRange("K2");
    teamLeaderContributionDueDateMinusOrdinandPlusYearAsReferenceForConditionalFormatting.setValue(
      teamLeaderContributionDueDateMinusOrdinandPlusYearAsReferenceForConditionalFormatting.getValue()
      .toString()
      .replace(/TEAM_LEADER_CONTRIBUTION_DUE_DATE_MINUS_ORDINAND_PLUS_YEAR_AS_REFERENCE_FOR_CONDITIONAL_FORMATTING/g,
        teamLeaderContributionDueDateMinusOrdinand + " " + currentYear));
    
    var staffSheetValues = staffSheet.getDataRange()
      .getValues();
    staffSheetValues = staffSheetValues.slice(staffSheet.getFrozenRows()); //remove headers if any
    var staff_array = staffSheetValues.map(function(c, i, a) {
      return {
        current_value_is_staff_first_name: c[0],
        current_value_is_team_name: c[11],
        current_value_is_team_leader: (c[12].toLowerCase() === "yes"),
      };
    }, []);
    var team_leaders_array = staff_array.filter(function(i) {
      return i.current_value_is_team_leader
    }); //remove non-team leaders
    for (var team_leaders_array_index = 0; team_leaders_array_index <
      team_leaders_array.length; team_leaders_array_index++) {
      var current_team_name = team_leaders_array[team_leaders_array_index].current_value_is_team_name;
      if (current_team_name == teamName) {
        var current_leader_first_name = team_leaders_array[
          team_leaders_array_index].current_value_is_staff_first_name;
        var please_complete = newTeamSheet.getRange("L1");
        please_complete.setValue(please_complete.getValue()
          .toString()
          .replace(/TEAM_LEADERS_CONTRIBUTION_DUE_DATE/g,
            teamLeadersContributionDueDate));
        please_complete.setValue(please_complete.getValue()
          .toString()
          .replace(/TEAM_LEADER_FIRST_NAME/g, current_leader_first_name));
      }
    }
    var startRow = 4;
    var startColumn = 1;
    var numberOfImportedRows = newTeamSheet.getLastRow() - 3;
    var firstRowNumberAfterImportedRows = newTeamSheet.getLastRow() + 1;
    var numberOfCurrentYearColumns = 2;
    var numberOfEventNameColumns = 1;
    var numberOfNullCellsRows = newTeamSheet.getMaxRows() -
      firstRowNumberAfterImportedRows;
    //set data validations
    var dataValidationRangeNo = newTeamSheet.getRange(startRow, 3,
      numberOfImportedRows, 1);
    var dataValidationRuleNo = newTeamSheet.getRange("C4")
      .getDataValidation();
    dataValidationRangeNo.setDataValidation(dataValidationRuleNo);
    var dataValidationRangeYes = newTeamSheet.getRange(startRow, 5,
      numberOfImportedRows, 1);
    var dataValidationRuleYes = newTeamSheet.getRange("E4")
      .getDataValidation();
    dataValidationRangeYes.setDataValidation(dataValidationRuleYes);
    //copy template data
    var cellsRangeNo = newTeamSheet.getRange(startRow, 4, numberOfImportedRows,
      1);
    cellsRangeNo.setValue("No");
    var cellsRangeYes = newTeamSheet.getRange(startRow, 6, numberOfImportedRows,
      1);
    cellsRangeYes.setValue("Yes");
    //protect headers, imported data, template data, and null cells
    var currentYearRange = newTeamSheet.getRange(startRow, 1,
      numberOfImportedRows, numberOfCurrentYearColumns);
    currentYearRange.setBackground("#efefef")
      .protect()
      .setDescription(newTeamSheet.getName() + ": " + (currentYear) + " " +
        "Promo Tiers and Start Dates");
    var importedEventsRangeUnprotected = newTeamSheet.getRange(startRow, 7,
      numberOfImportedRows, 4);
    importedEventsRangeUnprotected.setBorder(null, null, true, null, null, null,
      "#163c47", SpreadsheetApp.BorderStyle.DASHED)
    var importedEventsRangeProtected = newTeamSheet.getRange(startRow, 7,
      numberOfImportedRows, 1);
    importedEventsRangeProtected.protect()
      .setDescription(newTeamSheet.getName() + ": " + (currentYear) + " " +
        "Events");
    var nullCellsRange = newTeamSheet.getRange(firstRowNumberAfterImportedRows,
      1, numberOfNullCellsRows + 1, 6);
    nullCellsRange.merge()
      .setBorder(true, null, null, true, null, null, "#000000", SpreadsheetApp.BorderStyle
        .SOLID_MEDIUM)
      .protect()
      .setDescription(newTeamSheet.getName() + ": " + "Null Cells");
    var headerRange = newTeamSheet.getRange(1, 1, 3, 13);
    headerRange.protect()
      .setDescription(newTeamSheet.getName() + ": " + "Headers");
    var noRange = newTeamSheet.getRange(startRow, 4, numberOfImportedRows, 1);
    noRange.protect()
      .setDescription(newTeamSheet.getName() + ": " + "No");
    var yesRange = newTeamSheet.getRange(startRow, 6, numberOfImportedRows, 1);
    yesRange.protect()
      .setDescription(newTeamSheet.getName() + ": " + "Yes");
    //set protection edit permissions only to script executor 
    var me = Session.getEffectiveUser();
    var protections = newTeamSheet.getProtections(SpreadsheetApp.ProtectionType
      .RANGE);
    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.canEdit()) {
        protection.addEditor(me);
        protection.removeEditors(protection.getEditors());
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }
      }
    }
  }
}
// Private Functions
// -----------------
function nth_(ordinand) {
  if (ordinand > 3 && ordinand < 21) return 'th';
  switch (ordinand % 10) {
    case 1:
      return "st";
    case 2:
      return "nd";
    case 3:
      return "rd";
    default:
      return "th";
  }
}

function getWeeksInMonth_(month, year) {
  var year = new Date()
    .getYear();
  var month = 9;
  var weeks_array = [],
    firstDate = new Date(year, month, 1),
    lastDate = new Date(year, month + 1, 0),
    numDays = lastDate.getDate();
  var start = 1;
  var end = 7 - firstDate.getDay();
  while (start <= numDays) {
    weeks_array.push({
      start: start,
      end: end
    });
    start = end + 1;
    end = end + 7;
    end = start === 1 && end === 8 ? 1 : end;
    if (end > numDays)
      end = numDays;
  }
  return weeks_array;
}

function getFirstWeek_() {
  var weeks_array = getWeeksInMonth_();
  var year = new Date()
    .getYear();
  var month = 9;
  var m_names = ['January', 'February', 'March',
    'April', 'May', 'June',
    'July', 'August', 'September',
    'October', 'November', 'December'
  ];
  d = new Date(year, month, 1, 0, 0, 0, 0);
  var n = m_names[d.getMonth()];
  for (var weeks_array_index = 0; weeks_array_index < weeks_array.length; weeks_array_index++) {
    var current_start = weeks_array[weeks_array_index].start;
    var current_end = weeks_array[weeks_array_index].end;
    var first_Friday = current_end - 1;
    var teamLeadersContributionDueDate = "Friday, " + n + " " + first_Friday + nth_(current_end);
    // var teamLeadersContributionDueDate_minus_nth = "Friday, " + n + " " + first_Friday);
    if (current_end - current_start == 6) {
      return teamLeadersContributionDueDate;
    }
  }
}