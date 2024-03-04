let startAt = 0;
let currentDateValue = new Date();
// Calculate start date and end date dynamically based on the current date
let oneDayBeforeVal = new Date(currentDateValue);
oneDayBeforeVal.setDate(oneDayBeforeVal.getDate() - 1);

let oneMonthBeforeVal = new Date(oneDayBeforeVal);
oneMonthBeforeVal.setMonth(oneMonthBeforeVal.getMonth() - 1);

let formattedStartDateVal = new Date(oneMonthBeforeVal);
let formattedEndDateVal = new Date(oneDayBeforeVal);

let startDateVal = getFormattedDate(formattedStartDateVal);
let endDateVal = getFormattedDate(formattedEndDateVal);
Logger.log(endDateVal);
function getLeaveReport() {
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
  // Logger.log(attendanceData);
  Logger.log(apiUrl);
  let headers = [
    ["Employee Name", "Status", "Date"]
  ];
  let finalValues = [];
  attendanceData.forEach(record => {
    let empDetails = record.employeeDetails;
    let attendanceRecord = record.attendanceDetails;
    let sortedKeys = Object.keys(attendanceRecord).sort();
    sortedKeys.forEach(key => {
      let firstIn = getTimeFormat(attendanceRecord[key]['FirstIn']);
      let lastOut = getTimeFormat(attendanceRecord[key]['LastOut']);
      let netTimeDifference = calculateNetTimeDifference(attendanceRecord[key]['DeviationTime'], attendanceRecord[key]['TotalHours']);
      let rowArray = [
        `${empDetails['first name']} ${empDetails['last name']}`,
        attendanceRecord[key]['Status'],
        new Date(key).toLocaleDateString(),
      ];
      finalValues.push(rowArray);
    });
  });

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave Report(Zoho)");
  let lastRow = sheet.getLastRow();
  let increaseLimit = 1;
  if (lastRow == 0) {
    increaseLimit = 2;
    sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  }
  if(finalValues.length > 0)
  sheet.getRange(lastRow + increaseLimit, 1, finalValues.length, finalValues[0].length).setValues(finalValues);

  // If all records are not fetched then again call the API until all records are fetched.
  if(attendanceData.length){
    startAt = startAt + 100;
    getLeaveReport();
  }
}
