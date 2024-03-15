// Function for generating attendance report
function getAttendanceStatusReport(month=2) {
    givenMonth = (month) ? month : givenMonth;
    let datesObj = getFirstAndLastDateOfMonth(givenMonth)
    Logger.log(datesObj);
    let startDate = datesObj.firstDate;
    let endDate = datesObj.lastDate;
    let monthVal = datesObj.month;
    let sheetName = `${getSheetNameByMonth(monthVal)} Attendance details`
    let attendanceReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!attendanceReportSheet) {
      // If the sheet doesn't exist, create a new one
      attendanceReportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      Logger.log("Created new sheet: " + sheetName);
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
      ["Employee Id", "Employee Name", "Email ID", "Date", "First In", "Last Out", "Total Hours", "Early Entry", "Late Entry", "Early Exit", "Late Exit", "Net hours", "Shift Name", "Status", "Worked", "Paid", "Unpaid"]
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
          let workedDaysCount = 0;
          let paidDays = 0;
          let unpaidDays = 0;
          if( attendanceRecord[key]['Status'] == "Present" || attendanceRecord[key]['Status'].includes("0.5 day Present")){
            if(attendanceRecord[key]['LeaveCode']){
              if(attendanceRecord[key]['LeaveCode'].includes("0.5 day")){
                workedDaysCount = 0.5;
                paidDays = 0.5
              } else {
                workedDaysCount = 1; 
                paidDays = 1
              }
            } else {
              workedDaysCount = 1; 
              paidDays = 1
            }
          }
          if( attendanceRecord[key]['Status'].includes("Weekend") || attendanceRecord[key]['Status'].includes("Holiday")){
            paidDays = 1;
          }
          if( attendanceRecord[key]['Status'] == "Absent"){
            unpaidDays = 1;
          }
          if( attendanceRecord[key]['Status'].includes("Leave") ){
            let availableBalance = getLeaveBalance(empDetails.erecno,attendanceRecord[key]['Status']);
            if(availableBalance){
              paidDays = 1;
            } else {
              unpaidDays = 1;
            }
          }
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
            attendanceRecord[key]['Status'],
            workedDaysCount,
            paidDays,
            unpaidDays
          ];
          finalValues.push(rowArray);
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
      getAttendanceStatusReport();
    }
  }