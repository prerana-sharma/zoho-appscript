// Function to parse time string and convert it to minutes
function parseTimeToMinutes(timeString) {
  let [hours, minutes] = timeString.split(':').map(Number);
  return hours * 60 + minutes;
}

// Function to calculate net time difference from deviation time and total hours
function calculateNetTimeDifference(deviationTime, totalHours) {
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
// Function for getting date in "YYYY-MM-DD" format
function getFormattedDate(dateString){
  let currentDate = new Date(dateString);
  let year = currentDate.getFullYear();
  let month = String(currentDate.getMonth() + 1).padStart(2, '0'); // Adding 1 because months are zero-indexed
  let day = String(currentDate.getDate()).padStart(2, '0');

  // Construct the date string in "YYYY-MM-DD" format
  let formattedDate = `${year}-${month}-${day}`;
  return formattedDate;
}

function getFirstAndLastDateOfMonth(month) {
  if(!month){
    let currentDateValue = new Date();
    // Get the current month 
    let currentMonth = currentDate.getMonth() + 1;
    month = currentMonth;
  }
  let currentYear = new Date().getFullYear(); // Get the current year

  // Initialize variables to hold the first and last dates
  let firstDate, lastDate;

  // Use a switch statement to handle each month
  switch (month) {
      case 1: // January
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 31);
          break;
      case 2: // February
          // Check if it's a leap year
          if ((currentYear % 4 == 0 && currentYear % 100 != 0) || currentYear % 400 == 0) {
            // Leap year
            firstDate = new Date(currentYear, month - 1, 1);
            lastDate = new Date(currentYear, month - 1, 29);
          } else {
            // Non-leap year
            firstDate = new Date(currentYear, month - 1, 1);
            lastDate = new Date(currentYear, month - 1, 28);
          }
          break;
      case 3: // March
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 31);
          break;
      case 4: // April
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 30);
          break;
      case 4: // May
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 31);
          break;
      case 5: // June
          firstDate = new Date(year, month - 1, 1);
          lastDate = new Date(year, month - 1, 30);
          break;
      case 4: // July
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 31);
          break;
      case 4: // August
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 31);
          break;
      case 4: // September
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 30);
          break;
      case 4: // October
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 31);
          break;
      case 4: // November
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 30);
          break;
      case 4: // December
          firstDate = new Date(currentYear, month - 1, 1);
          lastDate = new Date(currentYear, month - 1, 31);
          break;
      default:
          // Handle invalid month
          console.error("Invalid month number. Please provide a number between 1 and 12.");
          return;
  }
  // Format the dates if needed
  let formattedFirstDate = getFormattedDate(firstDate);
  let formattedLastDate = getFormattedDate(lastDate);

  return {
      firstDate: formattedFirstDate,
      lastDate: formattedLastDate,
      month: month
  };
}

function getSheetNameByMonth(month){
  switch (month) {
    case 1:
      return "January";
    case 2:
      return "February";
    case 3:
      return "March";
    case 4:
      return "April";
    case 5:
      return "May";
    case 6:
      return "June";
    case 7:
      return "July";
    case 8:
      return "August";
    case 9:
      return "September";
    case 10:
      return "October";
    case 11:
      return "November";
    case 12:
      return "December";
    default:
      return "Invalid month";
  }
}

function clearSheetRows(startTimeColValues, dateThreshold, attendanceReportSheet){
  if(startTimeColValues[0].length){
      // Identify rows to delete
      let rowsToDelete = [];
      for (let i = startTimeColValues.length - 1; i >= 1; --i) {
        let sheetDate = new Date(startTimeColValues[i][0]).toLocaleDateString();
        if (new Date(sheetDate) >= dateThreshold) {
          rowsToDelete.push(i+1); // Push row numbers (1-based) to delete
        }
      }
      // Delete rows in batches
      const batchSize = 200;
      for (let i = 0; i < rowsToDelete.length; i += batchSize) {
        let batch = rowsToDelete.slice(i, i + batchSize);
        attendanceReportSheet.deleteRows(batch[batch.length -1], batch.length);
      }
    }
}
