let startAt = 0;
let todayDateValue = new Date();
let givenMonth = todayDateValue.getMonth() + 1;
let alreadyCall = false;
// Function for generating leave report
function getLeaveReport(month) {
  givenMonth = (month) ? month : givenMonth;
  let datesObj = getFirstAndLastDateOfMonth(givenMonth)
  Logger.log(datesObj);
  let startDateVal = datesObj.firstDate;
  let endDateVal = datesObj.lastDate;
  let monthVal = datesObj.month;
  let leaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${getSheetNameByMonth(monthVal)} Leave`);
  if (!leaveSheet) {
    // If the sheet doesn't exist, create a new one
    leaveSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(`${getSheetNameByMonth(monthVal)} Leave`);
    Logger.log("Created new sheet: " + `${getSheetNameByMonth(monthVal)} Leave`);
  }
  if(!alreadyCall){
    let formattedStartDateVal = new Date(startDateVal).toLocaleDateString();
    const startTimeColValues = leaveSheet.getRange("C:C").getValues();
    const dateThreshold = new Date(formattedStartDateVal);
    clearSheetRows(startTimeColValues, dateThreshold, leaveSheet);
    alreadyCall = true;
  }
  let apiUrl = `https://people.zoho.com/people/api/attendance/getUserReport?sdate=${startDateVal}&edate=${endDateVal}&dateFormat=yyyy-MM-dd&startIndex=${startAt}`;
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
    ["Employee Name", "Status", "Date"]
  ];
  let finalValues = [];
  attendanceData.forEach(record => {
    let empDetails = record.employeeDetails;
    let attendanceRecord = record.attendanceDetails;
    let sortedKeys = Object.keys(attendanceRecord).sort();
    sortedKeys.forEach(key => {
      let rowArray = [
        `${empDetails['first name']} ${empDetails['last name']}`,
        attendanceRecord[key]['Status'],
        new Date(key).toLocaleDateString(),
      ];
      finalValues.push(rowArray);
    });
  });
  let lastRow = leaveSheet.getLastRow();
  let increaseLimit = 1;
  if (lastRow == 0) {
    increaseLimit = 2;
    leaveSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  }
  if(finalValues.length > 0)
  leaveSheet.getRange(lastRow + increaseLimit, 1, finalValues.length, finalValues[0].length).setValues(finalValues);

  // If all records are not fetched then again call the API until all records are fetched.
  if(attendanceData.length){
    startAt = startAt + 100;
    getLeaveReport();
  }
}
