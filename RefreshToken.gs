let scriptProperties = PropertiesService.getScriptProperties();
const REFRESH_TOKEN = scriptProperties.getProperty("refreshToken");
const CLIENT_ID = scriptProperties.getProperty("clientId");
const CLIENT_SECRET = scriptProperties.getProperty("clientSecret");
function getAccessToken() {
  let apiUrl = `https://accounts.zoho.com/oauth/v2/token?refresh_token=${REFRESH_TOKEN}&client_id=${CLIENT_ID}&client_secret=${CLIENT_SECRET}&grant_type=refresh_token`;
  let options = {
    "method": 'post',
    'muteHttpExceptions': true
  }; 
  let response = UrlFetchApp.fetch(apiUrl, options);
  let results = JSON.parse(response);
  Logger.log(results);
  scriptProperties.setProperty("accessToken", results.access_token);
}

function getTimeFormat(timeString="-") {
  let time = timeString.split(' ');
  if(time.length > 2) return `${time[1]} ${time[2]}`;
  return false;
}
