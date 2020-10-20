var COURSE_NAME = "CSCI";
var COURSE_ID = "164804044482";
const userKey = "all";
const applicationName = "meet";
var ui = SpreadsheetApp.getUi();
var timeZone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

var documentProperties = PropertiesService.getDocumentProperties();
var sessionLength = documentProperties.getProperty("SESSION_LENGTH");
var lateTolerance = documentProperties.getProperty("LATE_TOLERANCE");
var earlyTolerance = documentProperties.getProperty("EARLY_TOLERANCE");
var checkLateEarly = documentProperties.getProperty("CHECK_LATE_EARLY");

function onOpen () {
  if (sessionLength == null) documentProperties.setProperty("SESSION_LENGTH", "30");
  if (lateTolerance == null) documentProperties.setProperty("LATE_TOLERANCE", "3");
  if (earlyTolerance == null) documentProperties.setProperty("EARLY_TOLERANCE", "2");
  if (checkLateEarly == null) documentProperties.setProperty("CHECK_LATE_EARLY", "true");
  ui.createMenu("Automated Meet Attendance")
    .addItem("Setup classes", "showCourseSelector")
    .addSubMenu(ui.createMenu("Add a column for today")
      .addItem("On this sheet", "addToday")
      .addItem("On all sheets", "addTodayAllSheets"))
    .addSubMenu(ui.createMenu("Check attendance")
      .addItem("On this sheet", "checkAll")
      .addItem("On all sheets", "checkAllSheets"))
    .addSubMenu(ui.createMenu("Set usage properties")
      .addItem("Session length", 'promptSessionLength')
      .addSeparator()
      .addItem("Flag lateness/early departure", 'promptCheckLateEarly')
      .addItem("Late entry flag time", 'promptLateTolerance')
      .addItem("Early departure flag time", 'promptEarlyTolerance'))
    .addToUi();
}

function promptSessionLength () {
  while (true) {
    var response = ui.prompt("Set Session Length","How long, in minutes, do your class sessions last?",ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      var newSessionLength = new Number(response.getResponseText());
      if (newSessionLength != NaN) {
        documentProperties.setProperty("SESSION_LENGTH", newSessionLength);
        ui.alert("Session length set to " + newSessionLength + " minutes.");
        break;
      } else {
        ui.alert("Please enter a number.")
      }
    } else {
      break;
    }
  }
}

function promptLateTolerance () {
  while (true) {
    var response = ui.prompt("Set Tardiness Tolerance","How long, in minutes, should a user be excused when showing up after the session start time?",ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      var newLateTolerance = new Number(response.getResponseText());
      if (newLateTolerance != NaN) {
        documentProperties.setProperty("LATE_TOLERANCE", newLateTolerance);
        ui.alert("Tardiness tolerance set to " + newLateTolerance + " minutes.");
        break;
      } else {
        ui.alert("Please enter a number.")
      }
    } else {
      break;
    }
  }
}

function promptEarlyTolerance () {
  while (true) {
    var response = ui.prompt("Set Early Departure Tolerance","How long, in minutes, should a user be excused if leaving before the session ends?",ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
      var newEarlyTolerance = new Number(response.getResponseText());
      if (newEarlyTolerance != NaN) {
        documentProperties.setProperty("EARLY_TOLERANCE", newEarlyTolerance);
        ui.alert("Early departure tolerance set to " + newEarlyTolerance + " minutes.");
        break;
      } else {
        ui.alert("Please enter a number.")
      }
    } else {
      break;
    }
  }
}

function promptCheckLateEarly () {
  var response = ui.alert("Check for Lateness/Early Departure","Do you want to automatically flag (underline) students who are late or leave early?",ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    documentProperties.setProperty("CHECK_LATE_EARLY", "true");
  } else if (response == ui.Button.NO) {
    documentProperties.setProperty("CHECK_LATE_EARLY", "false");
  }
}

function showCourseSelector() {
  ui.showModalDialog(HtmlService.createTemplateFromFile("Courses.html").evaluate(), "Course Selector");
}

/*
  Description: Option for teachers to import their courses
*/
function importCourses(selection) {
  var optionalArgs = {
    teacherId: 'me',
    courseStates: 'ACTIVE'
  };
  var courses = Classroom.Courses.list(optionalArgs).courses;
  for (var i = 0; i < courses.length; i++) {
    if (selection.includes(courses[i].id)) insertCourse(courses[i].name, courses[i].id);
  }
}

function processCourseSelections(formObject) {
//  var selection = [];
//  for (var key in formObject) {
//    selection.push(key);
//  }
  importCourses(formObject.keys());
}

/*
  Description: Create the Sheet for Course
  @param {String} courseName - Name of Course
  @param {String} courseId - Corresponding Classroom ID
*/
function insertCourse(courseName, courseId) {
  if (courseName == null) courseName = COURSE_NAME;
  if (courseId == null) courseId = COURSE_ID;
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = activeSpreadsheet.getSheetByName(courseName);
  
  if (newSheet != null) {
    return;
  }
  var today = new Date();
  newSheet = activeSpreadsheet.insertSheet();
  newSheet.setName(courseName);
  newSheet.getRange("A:Z").setFontFamily("Open Sans");
  newSheet.getRange("A1:D1").setFontFamily("Patua One").setFontSize(14).setBackground("#274e77").setFontColor("white");
  newSheet.getRange("D:D").setNumberFormat("#.0").setHorizontalAlignment("center");
  newSheet.getRange("A2:D3").setFontWeight("bold").setHorizontalAlignment("center");
  newSheet.getRange("D1").insertCheckboxes();
  newSheet.appendRow([courseName,'Replace with Main Meet Code']);
  newSheet.appendRow(["Breakout Room Links if Applicable"]);
  newSheet.getRange("A3:C3").mergeAcross().setHorizontalAlignment("center").setNote("Place breakout room links into the note (Right click → Insert note) for a date, separated by commas or on separate lines within the note.");
  newSheet.appendRow(['Last Name', 'First Name', 'Email Address', today]).setColumnWidth(4, 48).setRowHeight(3, 125);
  newSheet.deleteColumns(newSheet.getLastColumn()+1, newSheet.getMaxColumns()-newSheet.getLastColumn());
  newSheet.getRange("D2").setNumberFormat("yyy.mm.dd hh:mm ddd").setTextRotation(90);
  newSheet.setFrozenRows(2);
  newSheet.setFrozenColumns(3);
  var roster = getRoster(courseId);
  var studentLastNames = roster["studentLastNames"];
  var studentFirstNames = roster["studentFirstNames"];
  var studentEmails = roster["studentEmails"];
  for (var i = 0; i < studentLastNames.length; i++) {
    newSheet.appendRow([studentLastNames[i],studentFirstNames[i],studentEmails[i]]);
  }
  newSheet.getRange("A3:D").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  newSheet.getRange("A2:D2").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("center");
  newSheet.sort(1);
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint("#f9cb9c")
    .setGradientMidpointWithValue("#b6d7a8",SpreadsheetApp.InterpolationType.NUMBER,"30")
    .setGradientMaxpoint("#a4c2f4")
    .setRanges([newSheet.getRange("D3:D")])
    .build();
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("A")
    .setBackground("#e6b8af")
    .setRanges([newSheet.getRange("D3:D")])
    .build();
  var rules = newSheet.getConditionalFormatRules();
  rules.push(rule1);
  rules.push(rule2);
  newSheet.setConditionalFormatRules(rules);
  newSheet.deleteRows(newSheet.getLastRow()+1, newSheet.getMaxRows()-newSheet.getLastRow());
}

/*
  Description: Adds the course's students to the course sheet
  @param {String} courseId - Corresponding Classroom ID
*/
function getRoster(courseId) {
  var studentLastNames = [];
  var studentFirstNames = [];
  var studentEmails = [];
  var optionalArgs = {};
  do {
    var response = Classroom.Courses.Students.list(courseId, optionalArgs);
    optionalArgs = { pageToken: response.nextPageToken };
    var students = response.students;
    
    for (var i = 0; i <= students.length; i++) {
      try {
        studentLastNames.push(students[i].profile.name.familyName);
        studentFirstNames.push(students[i].profile.name.givenName);
        studentEmails.push(students[i].profile.emailAddress);
      } catch (err) {
        continue;
      }
    }
  } while (response.nextPageToken != null);
  return { "studentLastNames":studentLastNames, "studentFirstNames":studentFirstNames, "studentEmails":studentEmails };
}

function addTodayAllSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0; i<sheets.length; i++) {
    addToday(sheets[i]);
  }
}

function addToday(sheet) {
  if (sheet == null) sheet = SpreadsheetApp.getActiveSheet();
  var previousDate = new Date(sheet.getRange("D2").getValue());
  Logger.log(sheet.getRange("D2").getValue());
  var today = new Date(Date.now());
  today.setHours(previousDate.getHours());
  today.setMinutes(previousDate.getMinutes());
  today.setSeconds(previousDate.getSeconds());
  today.setMilliseconds(previousDate.getMilliseconds());
  sheet.insertColumnBefore(4);
  sheet.getRange("D2:D3").mergeVertically();
  sheet.getRange("D2").setValue(today);
}


function checkAllSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0; i<sheets.length; i++) {
    SpreadsheetApp.setActiveSheet(sheets[i]);
    checkAll(sheets[i]);
  }
}

/*
  Description: Retrieves the Meet code from the Course Sheet
  and uses helper function to check attendance
*/
function checkAll(sheet) {
  if (sheet == null) sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var oldNotes = sheet.getDataRange().getNotes();
  var meetCodes = [getCleanCode(values[0][1])];
  // No Meet code given
  if (meetCodes == null) {
    return;
  }
  // First loop over the dates
  for (var j = 3; j < values[1].length; j++){
    // If the top row of the column is empty or checked, ignore it
    if (values[0][j] != "" && values[0][j] != null) continue;
    // Check for additional Meet links in the frozen header rows for the date
    try {
      var note = oldNotes[0][j] + oldNotes[1][j] + oldNotes[2][j];
      if (note != "" && note != null) meetCodes = meetCodes.concat(getCleanCodes(note));
    } catch (err) {}
    // Logger.log(meetCodes);
    var startTime = new Date(values[1][j]);
    startTime.setHours(0);
    startTime.setMinutes(0);
    startTime.setSeconds(0);
    var endTime = new Date(startTime.getTime()+24*60*60*1000);
    var sessionStart = new Date(values[1][j]);
    // Logger.log("Checking %s on %s", meetCodes, startTime);
    var notes = [];
    var lines = [];
    var texts = [];
    // Then loop over the students for that date
    for (var i = 3; i < values.length; i++) {
      var meetAttendance = checkMeet(meetCodes, startTime, endTime, sessionStart, values[i][2]);
      notes.push([meetAttendance.note]);
      lines.push([meetAttendance.line]);
      texts.push([meetAttendance.text]);
    }
    // Finally, update the spreadsheet for that date
    sheet.getRange(4, j+1, sheet.getMaxRows()-3).setNotes(notes).setFontLines(lines).setValues(texts);
    // And mark the attendance as taken
    sheet.getRange(1, j+1).check();
  }
}

/*
  Description: Checks the Meet for attendance of the given student
  @param {String} meetCodes - Raw Meet Code from Course Sheet
*/
function checkMeet(meetCodes, searchStartTime, searchEndTime, sessionStart, emailAddress) {
  var allActivitiesArray = [];
  for (var meetCounter = 0; meetCounter < meetCodes.length; meetCounter++) {
    var optionalArgs = {
      event_name: "call_ended",
      start_time: Utilities.formatDate(searchStartTime, timeZone, "yyyy-MM-dd'T'HH:mm:ss'Z'"),
      end_time: Utilities.formatDate(searchEndTime, timeZone, "yyyy-MM-dd'T'HH:mm:ss'Z'"),
      filters: "identifier==" + emailAddress + ",meeting_code==" + meetCodes[meetCounter]
    };
    try {
      var response = AdminReports.Activities.list(userKey, applicationName, optionalArgs);
      var activities = response.items;
      if (activities != null) {
        allActivitiesArray.push(activities);
      }
    } catch (err) {
      continue;
    }
  }
  // If there were no activities found, the student was absent
  if (allActivitiesArray.length == 0) {
    return {"note":"",
            "line":"none",
            "text":"A"};
  }
  var log = "";
  var flag = false;
  var simplifiedActivities = sortAndSimplifyActivities(allActivitiesArray);
  var includeMeetCode = (meetCodes.length != 1);
  for (var index = 0; index < simplifiedActivities.activities.length; index++) {
    var activity = simplifiedActivities.activities[index];
    var postfix = includeMeetCode?(" " + getDirtyCode(activity.meetCode).toLowerCase() + "\n"):("\n");
    log += Utilities.formatDate(activity["joinTime"], timeZone, "hh:mm a") + " joined Meet" + postfix;
    log += Utilities.formatDate(activity["exitTime"], timeZone, "hh:mm a") + " exited Meet" + postfix;
  }
  // Check for entering late or leaving early, only if desired by the user. If not, move on.
  if (checkLateEarly == "false") {
    return {"note":log,
            "line":"none",
            "text":simplifiedActivites.totalDuration/60};
  }
  // Check for joining late or leaving early.
  var firstJoinTime = new Date(simplifiedActivities.firstJoinActivity["joinTime"]).getTime();
  var lastExitTime = new Date(simplifiedActivities.lastExitActivity["exitTime"]).getTime();
  var sessionStartTime = sessionStart.getTime();
  var sessionEndTime = sessionStart.getTime() + Number(sessionLength)*60000;
  //Logger.log("firstJoinTime: %s\nlastExitTime: %s\nsessionStartTime: %s\nsessionEndTime: %s\n", new Date(firstJoinTime), new Date(lastExitTime), new Date(sessionStartTime), new Date(sessionEndTime));
  var flag = false;
  if (firstJoinTime - sessionStartTime > lateTolerance*60000) {
    log = "JOINED LATE\n" + log;
    flag = true;
  }
  if (sessionEndTime - lastExitTime > earlyTolerance*60000) {
    log = "LEFT EARLY\n" + log;
    flag = true;
  }
  if (lastExitTime < sessionStartTime || firstJoinTime > sessionEndTime) {
    log = "ABSENT\n" + log;
    flag = true;
  }
  return {"note":log,
          "line":(flag?"underline":"none"),
          "text":simplifiedActivities.totalDuration/60};
}

function getParameterValue(parameters, key) {
  for (var i = 0; i < parameters.length; i++) {
    if (parameters[i].name == key) return parameters[i];
  }
}

function sortAndSimplifyActivities(allActivitiesArray) {
  // List of all activities (once simplified)
  var activities = [];
  // Some useful metadata about the activities
  var firstJoinActivity;
  var lastExitActivity;
  var totalDuration = 0;
  // Merge any separate lists and simplify activities
  for (var i = 0; i < allActivitiesArray.length; i++) {
    for (var j = 0; j < allActivitiesArray[i].length; j++) {
      // Get information about the activity
      var parameters = allActivitiesArray[i][j].events[0].parameters;
      var meetCode = getParameterValue(parameters, "meeting_code").value;
      var duration = Number(getParameterValue(parameters, "duration_seconds").intValue);
      var exitTime = new Date(allActivitiesArray[i][j].id.time);
      var joinTime = new Date(exitTime.getTime()-(duration*1000));
      // Construct the activity
      var activity = { 
        "meetCode":meetCode,
        "joinTime":joinTime,
        "exitTime":exitTime,
        "duration":duration
      };
      // Insert the activity into the activities array, sorting by join time
      var insert = 0;
      while (insert < activities.length) {
        if (activities[insert]["joinTime"] > joinTime) break;
        insert++;
      }
      activities.splice(insert, 0, activity);
      totalDuration += duration;
    }
  }
  // Compare the activities pairwise to get the metadata
  for (var i = 0; i < activities.length; i++) {
    // Check (and update) if this activity was the first join or last exit
    if (firstJoinActivity == null || activities[i]["joinTime"] < firstJoinActivity["joinTime"]) firstJoinActivity = activities[i];
    if (lastExitActivity == null  || activities[i]["exitTime"] > lastExitActivity["exitTime"] )  lastExitActivity = activities[i];
    // If this activity overlaps with another (previous) activity (e.g., student used more than one device), subtract the overlap from the total duration
    for (var j = 0; j < i; j++) {
      if (activities[i]["joinTime"] < activities[j]["exitTime"]) {
        // Subtract the whole duration of this activity if it's wholly inside the other activity, or the time they overlap, in seconds
        totalDuration -= (Math.min(activities[i]["exitTime"],activities[j]["exitTime"]) - activities[i]["joinTime"])/1000;
      }
    }
  }
  return {"activities": activities,
          "firstJoinActivity": firstJoinActivity,
          "lastExitActivity": lastExitActivity,
          "totalDuration": totalDuration};
}

/*
  Description: Strips any "-' Characters to match needed format
  for Reports API
  @param {String} meetCodes - Raw Meet Code from Course Sheet
*/
function getCleanCode(meetCode) {
  try {
    if (meetCode.includes(".com")) {
      meetCode = meetCode.substring(meetCode.search(/.com\//i)+5);
    }
    return meetCode.replace(/-/g,"").trim();
  } catch (err) { 
    return meetCode;
  }
}

function getDirtyCode(meetCode) {
  return meetCode.substr(0,3) + "-" + meetCode.substr(3,4) + "-" + meetCode.substr(7);
}

function getCleanCodes(meetCodesString) {
  // Split at either commas or newline characters
  var meetCodes = meetCodesString.trim().split(/[,\n]+/g);
  var cleanCodes = [];
  for (var i = 0; i < meetCodes.length; i++) {
    try {
      cleanCodes.push(getCleanCode(meetCodes[i]));
    } catch (err) {
      Logger.log("Attempted unsucessfully to get clean code for " + meetCodes[i]);
    }
  }
  return cleanCodes;
}