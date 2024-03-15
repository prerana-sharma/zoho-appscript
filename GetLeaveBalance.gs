//Function to fetch the available leave balance of an employee
function getLeaveBalance(employeeId,leaveName) {
    var apiUrl = `https://people.zoho.com/people/api/v2/leavetracker/reports/user?employee=${employeeId}`;
    let options = {
      "method": 'get',
      'headers':{
        "Authorization" : `Zoho-oauthtoken ${ACCESS_TOKEN}`,
      },
      'muteHttpExceptions': true
    }; 
    try {
      let response = UrlFetchApp.fetch(apiUrl, options);
      let results = JSON.parse(response);
      if (results.leavetypes) {
        const record = results.leavetypes.find(record => leaveName.includes(record.leavetypeName));
        return record ? record.available : null;
      }
    } catch (error) {
      console.error("Error fetching epic data:", error);
    }
  
    return false;
  }
  