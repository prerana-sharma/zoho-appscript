const ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("accessToken");

let startLimit = 0;
let maxResultsLimit = 200;
let totalLimit = 0; 
let alreadyCallVal = false;

// Function for generating attendance report
function getAttendanceReport(month) {
  givenMonth = (month) ? month : givenMonth;
  let datesObj = getFirstAndLastDateOfMonth(givenMonth)
  Logger.log(datesObj);
  let startDate = datesObj.firstDate;
  let endDate = datesObj.lastDate;
  let monthVal = datesObj.month;
  let sheetNaem = `${getSheetNameByMonth(monthVal)} Attendance`
  let attendanceReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNaem);
  if (!attendanceReportSheet) {
    // If the sheet doesn't exist, create a new one
    attendanceReportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetNaem);
    Logger.log("Created new sheet: " + sheetNaem);
  }
  if(!alreadyCallVal){
    let formattedStartDateVal = new Date(startDate).toLocaleDateString();
    const startTimeColValues = attendanceReportSheet.getRange("D:D").getValues();
    const dateThreshold = new Date(formattedStartDateVal);
    clearSheetRows(startTimeColValues, dateThreshold, attendanceReportSheet);
    alreadyCallVal = true;
  }
  let apiUrl = `https://people.zoho.com/people/api/attendance/getUserReport?sdate=${startDate}&edate=${endDate}&dateFormat=yyyy-MM-dd&startIndex=${startLimit}`;
  let options = {
    "method": 'get',
    'headers':{
      "Authorization" : `Zoho-oauthtoken ${ACCESS_TOKEN}`,
    },
    'muteHttpExceptions': true
  }; 
  let response = UrlFetchApp.fetch(apiUrl, options);
  let results = JSON.parse(response);
  let attendanceData = results.result;
  let headers = [
    ["Employee Id", "Employee Name", "Email ID", "Date", "First In", "Last Out", "Total Hours", "Early Entry", "Late Entry", "Early Exit", "Late Exit", "Net hours", "Shift Name"]
  ];
  let finalValues = [];
  attendanceData.forEach(record => {
    let empDetails = record.employeeDetails;
    let attendanceRecord = record.attendanceDetails;
    let sortedKeys = Object.keys(attendanceRecord).sort();
    sortedKeys.forEach(key => {
      if( attendanceRecord[key]['Status'] == "Present"){
        let firstIn = getTimeFormat(attendanceRecord[key]['FirstIn']);
        let lastOut = getTimeFormat(attendanceRecord[key]['LastOut']);
        let netTimeDifference = calculateNetTimeDifference(attendanceRecord[key]['DeviationTime'], attendanceRecord[key]['TotalHours']);
        let rowArray = [
          empDetails.id,
          `${empDetails['first name']} ${empDetails['last name']}`,
          empDetails['mail id'],
          new Date(key).toLocaleDateString(),
          (firstIn) ? firstIn : "-",
          (lastOut) ? lastOut : "-",
          attendanceRecord[key]['TotalHours'],
          (attendanceRecord[key]['Early_In']) ? `+${attendanceRecord[key]['Early_In']}` : "-",
          (attendanceRecord[key]['Late_In']) ? `- ${attendanceRecord[key]['Late_In']}` : "-",
          (attendanceRecord[key]['Early_Out']) ? `- ${attendanceRecord[key]['Early_Out']}` : "-",
          (attendanceRecord[key]['Late_Out']) ? `+${attendanceRecord[key]['Late_Out']}` : "-",
          netTimeDifference,
          `[${attendanceRecord[key]['ShiftStartTime']} - ${attendanceRecord[key]['ShiftEndTime']}] ${attendanceRecord[key]['ShiftName']}`,
        ];
        finalValues.push(rowArray);
      }
    });
  });
  let lastRow = attendanceReportSheet.getLastRow();
  let increaseLimit = 1;
  if (lastRow == 0) {
    increaseLimit = 2;
    attendanceReportSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  }
  if(finalValues.length > 0)
  attendanceReportSheet.getRange(lastRow + increaseLimit, 1, finalValues.length, finalValues[0].length).setValues(finalValues);

  // If all records are not fetched then again call the API until all records are fetched.
  if(attendanceData.length){
    startLimit = startLimit + 100;
    getAttendanceReport();
  }
}

function createMenu() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom')
    .addItem('Fetch attendance record', 'showAttendanceDialog')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Fetch Leave record')
          .addItem('Select month', 'showCustomDialog'))
    .addToUi();
}
function onOpen(e) {
  createMenu();
}
function showCustomDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Leave');
}

function processArgument(argument) {
  getLeaveReport(parseInt(argument));
}
function showAttendanceDialog() {
  var html = HtmlService.createHtmlOutputFromFile('AttendanceDialog')
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Attendance');
}

function processAttendanceArgument(argument) {
  getAttendanceReport(parseInt(argument));
}