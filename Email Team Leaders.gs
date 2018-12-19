/*
***************************************************************************************

Redevelopment Notes from Chad: Should include another email soliciting help from Exec Director: 

//Hi {Executive Director}
//
//Please could I ask a favor? Would it be possible for you to send a word to each of the other team leaders on this? 
//
//Perhaps with the exception of Shawne, who is the only team leader to have contributed thus far.
//
//It could say something like what’s below (please revise appropriately).
//
//{Communications Director}
//
//
//— 
//
//Hi _, 
//
//Before we get into top gear with year-end activities and the pastoral transition process, I would like to encourage each of the team leaders to take a few moments to fill out their planning calendars for 2019. See the link below.
//
//If you need help on this, and/or will not be able to complete your calendar, by Friday, Oct 19—Please reply to let me know for forward a message to inform Chad.
//
//You have shepherded so many successful events, groups, and other creative campaigns for Christ Church in 2018. Every time we communicate what the life of the church is doing, and people participate, it means our community is formed more and more into the shape and image of the Kingdom of God in Nashville. Let’s follow God in this mission even further in the new year! 
//
//I so appreciate you,
//
//{Executive Director}

***************************************************************************************
*/

// email to Team Leaders, soliciting their contribution to the Promotion Planning Calendars for Teams by deadline
function inviteTeamLeaderstoContributebythefirstFridayOfOctober() {
  var staff_array = getStaffArray_(); //get all staff members
  var teamLeadersContributionDueDate = getFirstWeek_();
  var team_leaders_array = staff_array.filter(function(i) {
    return i.current_value_is_team_leader
  }); //remove non-team leaders
  for (var team_leaders_array_index = 0; team_leaders_array_index < team_leaders_array.length; team_leaders_array_index++) {
    var Spreadsheet = SpreadsheetApp.openById(PROMOTION_PLANNING_CALENDAR_FOR_TEAMS_ID_);
    var current_team_name = team_leaders_array[team_leaders_array_index].current_value_is_team_name;
    var current_team_sheet = Spreadsheet.getSheetByName(current_team_name);
    var current_leader_first_name = team_leaders_array[team_leaders_array_index].current_value_is_staff_first_name;
    var current_leader_email = team_leaders_array[team_leaders_array_index].current_value_is_staff_email;
    var current_leader_sheet_url = Spreadsheet.getUrl() + "#gid=" + current_team_sheet.getSheetId();
    var subject = "Promotion planning for your team";
    var body = Utilities.formatString("\
Dear %s: <br><br>\
It’s that time again! Please complete the instructions for <a href='%s'>Promotion Planning Calendars for Teams</a>, on behalf of all your team members, by <strong>%s</strong>.<br><br>\
Please reply to this email with any questions.<br><br>\
Thank you!",
                                      current_leader_first_name,
                                      Spreadsheet,
                                      teamLeadersContributionDueDate
                                     );
    MailApp.sendEmail({
      name: "communications@ccnash.org",
      to: current_leader_email,
      subject: subject,
      htmlBody: body
    });
  }
}

// reminder email to Team Leaders who have not entered ANY data yet
function emailTeamLeadersWhoHaveNotDoneAnything() {
  var staff_array = getStaffArray_(); //get all staff members
  var team_leaders_array = staff_array.filter(function(i) {
    return i.current_value_is_team_leader
  }); //remove non-team leaders
  for (var team_leaders_array_index = 0; team_leaders_array_index < team_leaders_array.length; team_leaders_array_index++) {
    var Spreadsheet = SpreadsheetApp.openById(PROMOTION_PLANNING_CALENDAR_FOR_TEAMS_ID_);
    var ui = SpreadsheetApp.getUi();
    var current_team_name = team_leaders_array[team_leaders_array_index].current_value_is_team_name;
    var current_team_sheet = Spreadsheet.getSheetByName(current_team_name);
    var current_leader_first_name = team_leaders_array[team_leaders_array_index].current_value_is_staff_first_name;
    var current_leader_email = team_leaders_array[team_leaders_array_index].current_value_is_staff_email;
    var current_leader_sheet_url = Spreadsheet.getUrl() + "#gid=" + current_team_sheet.getSheetId();
    var todays_date = new Date();
    var current_team_sheet_array = current_team_sheet.getRange("B:B")
    .getValues();
    var last_row = current_team_sheet.getMaxRows();
    var first_row_for_entering_next_year_events = current_team_sheet_array.length + 4;
    if (current_team_sheet.getRange(first_row_for_entering_next_year_events, 7, last_row, 1)
      .isBlank() && current_team_sheet.getRange(4, 8, last_row, 3)
      .isBlank()) {
        var wristSlap =  ui.alert('Notice!', 
        Utilities.formatString("You're about to email %s to ask why NOTHING has been contributed to \n hi or her planning calendar yet.\n\n You sure?\n\n", current_leader_first_name), 
               ui.ButtonSet.YES_NO);
          if(wristSlap!='YES') break;
        var subject = "Just checking in";
        var body = Utilities.formatString("\
Hi %s,<br><br>\
I noticed you haven't had a chance to make any changes or additions to <a href='%s'>your team's Promotion Planning Calendar</a> yet.<br><br>\
Is there anything I can do to help?",
                                          current_leader_first_name,
                                          current_leader_sheet_url
                                         );
        MailApp.sendEmail({
          name: "communications@ccnash.org",
          to: current_leader_email,
          subject: subject,
          htmlBody: body
        });
      }
  }
}

// email to all team leaders to remind them week of deadline
function remindAllTeamLeadersToContributeByThisFriday() {
  var staff_array = getStaffArray_(); //get all staff members
  var team_leaders_array = staff_array.filter(function(i) {
    return i.current_value_is_team_leader
  }); //remove non-team leaders
  for (var team_leaders_array_index = 0; team_leaders_array_index < team_leaders_array.length; team_leaders_array_index++) {
    var Spreadsheet = SpreadsheetApp.openById(PROMOTION_PLANNING_CALENDAR_FOR_TEAMS_ID_);
    var ui = SpreadsheetApp.getUi();
    var current_team_name = team_leaders_array[team_leaders_array_index].current_value_is_team_name;
    var current_team_sheet = Spreadsheet.getSheetByName(current_team_name);
    var current_leader_first_name = team_leaders_array[team_leaders_array_index].current_value_is_staff_first_name;
    var current_leader_email = team_leaders_array[team_leaders_array_index].current_value_is_staff_email;
    var current_leader_sheet_url = Spreadsheet.getUrl() + "#gid=" + current_team_sheet.getSheetId();
    var teamLeadersContributionDueDate = getFirstWeek_();
    var wristSlap =  ui.alert('Note!', 
        Utilities.formatString("You're about to encourage %s to contribute by this coming Friday.\n\n You sure?\n\n", current_leader_first_name), 
               ui.ButtonSet.YES_NO);
          if(wristSlap!='YES') continue;
    var subject = "Don't forget! Your team's promotion planning";
    var body = Utilities.formatString("\
Dear %s: <br><br>\
Don't forget, <a href='%s'>your team's Promotion Planning Calendar</a> is due by this coming %s.<br><br>\
Please let me know if you have any questions or concerns.<br><br>\
Thank you!",
                                      current_leader_first_name,
                                      current_leader_sheet_url,
                                      teamLeadersContributionDueDate
                                     );
    MailApp.sendEmail({
      name: "communications@ccnash.org",
      to: current_leader_email,
      subject: subject,
      htmlBody: body
    });
  }
}

// naugty and nice list
function emailNaughtyandNice() {
  var staff_array = getStaffArray_(); //get all staff members
  var team_leaders_array = staff_array.filter(function(i) {
    return i.current_value_is_team_leader
  });
  for (var team_leaders_array_index = 0; team_leaders_array_index < team_leaders_array.length; team_leaders_array_index++) {
    var Spreadsheet = SpreadsheetApp.openById(PROMOTION_PLANNING_CALENDAR_FOR_TEAMS_ID_);
    var ui = SpreadsheetApp.getUi();
    var current_team_name = team_leaders_array[team_leaders_array_index].current_value_is_team_name;
    var current_team_sheet = Spreadsheet.getSheetByName(current_team_name);
    var current_leader_first_name = team_leaders_array[team_leaders_array_index].current_value_is_staff_first_name;
    var current_leader_email = team_leaders_array[team_leaders_array_index].current_value_is_staff_email;
    var current_leader_sheet_url = Spreadsheet.getUrl() + "#gid=" + current_team_sheet.getSheetId();
    var next_year = new Date()
    .getFullYear() + 1;
    var teamLeadersContributionDueDate = getFirstWeek_();
    var Cvals = current_team_sheet.getRange("C1:C")
    .getValues();
    var Clast = Cvals.filter(String)
    .length;
    var current_team_sheet_array_Col_C = current_team_sheet.getRange("C4:C")
    .getValues();
    var current_team_sheet_array_Col_E = current_team_sheet.getRange("E4:E")
    .getValues();
    var current_team_sheet_array_Col_J = current_team_sheet.getRange("J4:J")
    .getValues();
    for (var current_team_sheet_array_Col_C_index = 0; current_team_sheet_array_Col_C_index < Clast + 1; current_team_sheet_array_Col_C_index++)
      if (current_team_sheet_array_Col_C[current_team_sheet_array_Col_C_index].toString() === "Yes" &&
          current_team_sheet_array_Col_E[current_team_sheet_array_Col_C_index].toString() === "No" &&
          current_team_sheet_array_Col_J[current_team_sheet_array_Col_C_index].toString() === "") {
          var wristSlap =  ui.alert('Hmm...', 
                                    Utilities.formatString("You're about to slap %s's wrist.\n\n You sure?\n\n", current_leader_first_name), 
                                    ui.ButtonSet.YES_NO);
          if(wristSlap!='YES') break;

            var array_Colors = current_team_sheet.getRange("G4:"+"G"+Clast).getBackgrounds();
            var array_eventNames = current_team_sheet.getRange("G4:"+"G"+Clast).getValues();
            Array.prototype.zip = function (arr) {
              return this.map(function (e, i) {
                return [e, arr[i]];
              })
            };
            var combinedArray_ColorsAndEventNames = array_eventNames.zip(array_Colors);
            var keys = ["event", "color"]; 
            var values = combinedArray_ColorsAndEventNames;
            var resultArray = [];
            for(var i=0; i<values.length; i++){
              var obj = {};
              for(var j=0; j<keys.length; j++){
                obj[keys[j]] = values[i][j];
              }
              resultArray.push(obj);
            }
            var yellowEvents = resultArray.filter(function (e) {
              return e.color == "#ffff00";
            });
            var keyArray = yellowEvents.map(function(item) { return item["event"]; });
            var stringArray = keyArray + ""
            var stringArrayBreaks = stringArray.split(",").join("<li>");
            var subject = "Follow up";
          var body = Utilities.formatString("\
%s,<br><br>\
I noticed you weren't able to finish <a href='%s'>your team's Promotion Planning Calendar</a>.<br><br> These events on your calendar are missing information:<br><ul><li>%s</li></ul>\
Please would you either revise these events (highlighted in <span style='background-color: #FFFF00'>yellow</span> on your calendar), email me an update, or <a target='_blank' href='https://calendly.com/chadbarlow/promo'>schedule a promotion planning meeting</a> with me so we can work this out together?\
",                          
                                            current_leader_first_name,
                                            current_leader_sheet_url,
                                            stringArrayBreaks
                                           
                                           );
          MailApp.sendEmail({
            name: "communications@ccnash.org",
            to: current_leader_email,
            subject: subject,
            htmlBody: body
          });
          
          break;
        }
  if (current_team_sheet_array_Col_C_index === Clast + 1) {
    var loveYa = ui.alert('Right on!', 
                          Utilities.formatString("You're about to praise to %s.\n\nYou sure?\n\n", current_leader_first_name), 
                          ui.ButtonSet.YES_NO);
    if(loveYa!='YES') break;      var subject = "Thank you!";
    var body = Utilities.formatString("\
Hey %s,<br><br>\
I noticed that you were one of the 'good ones' to completely finish <a href='%s'> your team's Promotion Planning Calendar</a> by the deadline last week, \
and I wanted to say thank you.<br><br>\
Your input is invaluable for shaping the comms strategy for %s. Your punctuality is priceless!<br><br>\
Chad",
                                      current_leader_first_name,
                                      current_leader_sheet_url,
                                      next_year
                                     );
    MailApp.sendEmail({
      name: "communications@ccnash.org",
      to: current_leader_email,
      subject: subject,
      htmlBody: body
    });
  }
}
}

//Private Functions
//=================

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
    if (current_end - current_start === 6) {
      return teamLeadersContributionDueDate;
    }
  }
}

function getStaffArray_() {
  var sheet = SpreadsheetApp.openById(STAFF_DATA_ID_); //"Staff Data" //1HEOWmNPo32uhR6N1XkviYiDM7KdAnaYycKDH9fz3OXE "Staff Data - Test"
  var values = sheet.getDataRange()
  .getValues();
  values = values.slice(sheet.getFrozenRows()); //remove headers if any
  var staff_array = values.map(function(c, i, a) {
    return {
      current_value_is_staff_first_name: c[0],
      current_value_is_staff_email: c[8],
      current_value_is_team_name: c[11],
      current_value_is_team_leader: (c[12].toLowerCase() === "yes"),
    };
  }, []);
  return staff_array;
}
