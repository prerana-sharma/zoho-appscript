// Function to parse time string and convert it to minutes
function parseTimeToMinutes(timeString) {
  let [hours, minutes] = timeString.split(':').map(Number);
  return hours * 60 + minutes;
}

// Function to calculate net time difference from deviation time and total hours
function calculateNetTimeDifference(deviationTime, totalHours) {
  Logger.log(totalHours)
  if(totalHours == "00:00"){
    return "-";
  }
  if(deviationTime && totalHours){
    let deviationMinutes = parseTimeToMinutes(deviationTime);
    let totalMinutes = parseTimeToMinutes(totalHours);

    let netMinutes = totalMinutes - deviationMinutes;
    let sign = netMinutes < 0 ? '-' : '+'; // Determine sign
    netMinutes = Math.abs(netMinutes); // Convert to positive value for calculation

    let netHours = Math.floor(netMinutes / 60);
    let remainingMinutes = netMinutes % 60;

    return `${sign} ${netHours.toString().padStart(2, '0')}:${remainingMinutes.toString().padStart(2, '0')}`;
  }
  return '-';
}

function getFormattedDate(dateString){
  let currentDate = new Date(dateString);
  let year = currentDate.getFullYear();
  let month = String(currentDate.getMonth() + 1).padStart(2, '0'); // Adding 1 because months are zero-indexed
  let day = String(currentDate.getDate()).padStart(2, '0');

  // Construct the date string in "YYYY-MM-DD" format
  let formattedDate = `${year}-${month}-${day}`;
  return formattedDate;
}
