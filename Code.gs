/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall (e) {
  onOpen(e);
}
/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Extract Timesheets Sidebar', 'openSidebar')
      .addToUi();
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function openSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate() // evaluate MUST come before setting the NATIVE mode
      .setSandboxMode(HtmlService.SandboxMode.NATIVE)
      .setTitle('Extract Timesheets')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(ui);
}
/**
 * Launch the calendars extraction
 * 
 * Function called by the sidebar with:
 *    selectdCal - the list of calendar IDs, 
 *    startDate - the start date (as a String) and
 *    endDate - the end date (as a String)
 */
function extractAll(selectedCal, startDate, endDate) {
  var start = new Date(startDate);
  var end = new Date(endDate);
  for (i = 0; i < selectedCal.length; i++) {
    extractFrom(selectedCal[i], start, end);
  }
}
/**
 * Launch the extraction from 1 calendar
 * 
 *    calendarID - The ID of the calendar to extract 
 *    startDate - the start date (as a Date) and
 *    endDate - the end date (as a Date)
 */
function extractFrom(calendarID, startDate, endDate) {
  // Extract calendar events for the selected period
  var cal = CalendarApp.getCalendarById(calendarID);
  var calName = cal.getName();
  var events = cal.getEvents(startDate, endDate);
  var owner = calName;
  
  // Select or create a page to store the data for the owner
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet = null;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() == owner) {
      sheet = sheets[i];
      break;
    }
  }
  if (sheet == null) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(owner, sheets.length);
  SpreadsheetApp.setActiveSheet(sheet);
  sheet.clearContents(); // Clear sheet content
  
  // Set the header
  var row = 1;
  var header = [["Owner", "Project", "Title", "Start", "End", "Duration", "Quarter", "Month"]];
  var range = sheet.getRange(row,1,1,8);
  range.setValues(header);
  SpreadsheetApp.flush();
  
  for (var i=0;i<events.length;i++) {
    var title = events[i].getTitle().trim();
    var project = "";
    if (title.charAt(0) != "#") {
      project = "## Unknown ##";
    } else {
      if (title.indexOf(" ") == -1) {
        project = title.substring(1);
        title = "";
      } else {
        project = title.substring(1, title.indexOf(" "));
        title = title.substring(title.indexOf(" "));
      }
    }
    var start = events[i].getStartTime();
    var end = events[i].getEndTime();
    var duration = (end - start) / 1000 / 60 / 60;
    
    var quarter = getQuarterFrom(start);
    var month = getMonthFrom(start);
    
    row = row + 1;
    var details = [[owner, project, title, start, end, duration, quarter, month]];
    var range = sheet.getRange(row,1,1,8);
    range.setValues(details);
  }
  SpreadsheetApp.flush();

}
/**
 * Return the quarter in the format YYYYQ#
 */
function getQuarterFrom(aDate) {
  var quarters = ["Q1", "Q1", "Q1", "Q2", "Q2", "Q2", "Q3", "Q3", "Q3", "Q4", "Q4", "Q4"];
  return aDate.getFullYear()+quarters[aDate.getMonth()];
}
/**
 * Return the month of a date in the format MMM-YYYY
 */
function getMonthFrom(aDate) {
  var months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  return months[aDate.getMonth()]+"-"+aDate.getFullYear();
}

function include(File) {
  return HtmlService.createTemplateFromFile(File).evaluate().getContent();
};
