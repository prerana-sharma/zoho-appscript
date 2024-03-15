// Function for generating attendance report
function getPresentDaysReport(month = 2) {
  givenMonth = (month) ? month : givenMonth;
  let datesObj = getFirstAndLastDateOfMonth(givenMonth)
  Logger.log(datesObj);
  let startDate = datesObj.firstDate;
  let endDate = datesObj.lastDate;
  let monthVal = datesObj.month;
  let currentYear = new Date().getFullYear();
  const daysInMonth = new Date(currentYear, monthVal, 0).getDate();
  let sheetname = `${getSheetNameByMonth(monthVal)} Present Days`
  let reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname);
  if (!reportSheet) {
    // If the sheet doesn't exist, create a new one
    reportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetname);
    Logger.log("Created new sheet: " + sheetname);
  }
  if(!alreadyCallVal){
    let formattedStartDateVal = new Date(startDate).toLocaleDateString();
    const startTimeColValues = reportSheet.getRange("D:D").getValues();
    const dateThreshold = new Date(formattedStartDateVal);
    clearSheetRows(startTimeColValues, dateThreshold, reportSheet);
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
    ["Employee Id", "Employee Name", "Email ID", "Expected Payable Day(s)", "Worked Days","PaidOffDays","TotalPayableDays","Expected Working Day(s)"]
  ];
  let finalValues = [];
  attendanceData.forEach(record => {
    let empDetails = record.employeeDetails;
    let attendanceRecord = record.attendanceDetails;
    let sortedKeys = Object.keys(attendanceRecord).sort();
    let workedDaysCount = 0;
    let paidOffDays = 0
    let exceptedWorkingDays = daysInMonth;
    sortedKeys.forEach(key => {
      if( attendanceRecord[key]['Status'] == "Present"){
        workedDaysCount += 1; 
      }
      if( attendanceRecord[key]['Status'] == "Weekend" || attendanceRecord[key]['Status'].includes("Holiday")){
        paidOffDays += 1;
        exceptedWorkingDays -= 1;
      }
    });
    let rowArray = [
      empDetails.id,
      `${empDetails['first name']} ${empDetails['last name']}`,
      empDetails['mail id'],
      daysInMonth,
      workedDaysCount,
      paidOffDays,
      workedDaysCount+paidOffDays,
      exceptedWorkingDays
    ];
    finalValues.push(rowArray)
  });
  let lastRow = reportSheet.getLastRow();
  let increaseLimit = 1;
  if (lastRow == 0) {
    increaseLimit = 2;
    reportSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  }
  if(finalValues.length > 0)
  reportSheet.getRange(lastRow + increaseLimit, 1, finalValues.length, finalValues[0].length).setValues(finalValues);

  // If all records are not fetched then again call the API until all records are fetched.
  if(attendanceData.length){
    startLimit = startLimit + 100;
    getPresentDaysReport();
  }
}