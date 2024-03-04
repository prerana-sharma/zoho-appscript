const ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("accessToken");

let startLimit = 0;
let maxResultsLimit = 200;
let totalLimit = 0; 

let currentDate = new Date();
// Calculate start date and end date dynamically based on the current date
let oneDayBefore = new Date(currentDate);
oneDayBefore.setDate(oneDayBefore.getDate() - 1);

let oneMonthBefore = new Date(oneDayBefore);
oneMonthBefore.setMonth(oneMonthBefore.getMonth() - 1);

let formattedStartDate = new Date(oneMonthBefore);
let formattedEndDate = new Date(oneDayBefore);

let startDate = getFormattedDate(formattedStartDate);
let endDate = getFormattedDate(formattedEndDate);
Logger.log(endDate);
function getAttendanceReport() {
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
  // Logger.log(attendanceData);
  Logger.log(apiUrl);
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
          "-",
          `[${attendanceRecord[key]['ShiftStartTime']} - ${attendanceRecord[key]['ShiftEndTime']}] ${attendanceRecord[key]['ShiftName']}`,
        ];
        finalValues.push(rowArray);
      }
    });
  });

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Automated attendance");
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
    startLimit = startLimit + 100;
    getAttendanceReport();
  }
}
