function main(){

  var customDate = '';
  var date = customDate !== "" ? customDate : formatDate(new Date()); // Today's date
  // Logger.log(typeof date)  
  var lastValidRow;
  var columns = {
  dateColumn: 0,         // Column A
  marketColumn: 1,       // Column B
  testerColumn: 3,       // Column D
  hexNameColumn: 5,      // Column F
  taskColumn: 6,         // Column G
  startTimeColumn: 7,    // Column H
  endTimeColumn: 8,       // Column I
  comments : 15,           // Column P 
  // avgHexCompletedThisWeek : 16 // Column Q
  };

  // Email configuration
  var to = ['DOC@dish.com','david.bentolila@dish.com','shreyas.rane@dish.com', 'jeremiah.watson@dish.com'];
  var cc = ['dhruval.potla@gcbservices.com', 'vignesh.kumar@gcbservices.com','aman.sharma@gcbservices.com','gyan.pandey@gcbservices.com'];
  var subject = 'GCB Daily Summary - FCC - ' + date;
  var name = 'Dhruval Potla';


  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var unfilteredHeadings = data[0];
  // Filter headings array to include only those present in columns object
  var headings = unfilteredHeadings.filter(function(heading) {return Object.values(columns).includes(unfilteredHeadings.indexOf(heading));});
  // Logger.log(headings);

  data = data.slice(1);
  // Logger.log(data);

  // Loop through the keys in the 'columns' object to create indices for table columns
  var tableColumns = {}
  for (var key in columns) {
    // Update the value of each key with its index
    tableColumns[key] =  parseInt(Object.keys(columns).indexOf(key));
  }
  // Logger.log(tableColumns);
  // Logger.log(tableColumns.hexNameColumn);

  var combinedDateDataset = getCombinedDateData(data, date, columns);
  // Logger.log(combinedDateDataset.length);

  var hexCompletedToday = [...(new Set(combinedDateDataset.map(row => row[tableColumns.hexNameColumn])))].length;
  // Logger.log("No. of Hex completed: " + hexCompletedToday);

  var testersToday = (new Set(combinedDateDataset.map(row => row[tableColumns.testerColumn]))).size;
  // Logger.log("Total testers in the market: " + testersToday);

  // Logger.log(typeof combinedDateDataset[0][columns.taskColumn]);
  var hexDataset = getHexStatusData(combinedDateDataset, tableColumns);
  // Logger.log("Length of Hex Status Data: " + hexDataset.length);

  var idleTSDataset = getIdleTSData(combinedDateDataset, tableColumns);
  // Logger.log("Length of Trobleshoot/Idle Status Data: " + troubleIdleDataset.length);
  var idleTSDatasethours = getIdleTSDatahours(combinedDateDataset,tableColumns);

  var idleTSHoursToday = getIdleTSHoursToday(idleTSDatasethours, tableColumns);
  var totalIdleTSHoursThisWeek = 0;
  var avgHexCompletedThisWeek = getAvgHexCompleted(data, date, columns);
  var html = prepareEmailBody(date, hexCompletedToday, idleTSHoursToday, totalIdleTSHoursThisWeek, testersToday, avgHexCompletedThisWeek);
  // Logger.log(html);
  var hexStatusTable = createHexStatusTable(hexDataset, html, headings);
  // Logger.log(hexStatusTable);
  // var troubleIdleTable = [];

  // var idleTSStatusTable = createIdleTSTable(idleTSDataset, html, headings);
  var idleTSStatusTable = createIdleTSTable(idleTSDataset, hexStatusTable, headings);
  // Logger.log(idleTSStatusTable);

  var finalEmailBody = includeEmailSignatures(idleTSStatusTable);
  Logger.log(finalEmailBody);

  // Commented out to prevent sending email
  sendEmail(to, cc, subject, finalEmailBody, name);
}

function getIdleTSHoursToday(idleTSDatasethours, tableColumns){
  var totalDuration = 0;
  var startTimeColumn = tableColumns.startTimeColumn;
  var endTimeColumn = tableColumns.endTimeColumn;

  for (var i = 0; i < idleTSDatasethours.length; i++){
    var duration = (idleTSDatasethours[i][endTimeColumn] - idleTSDatasethours[i][startTimeColumn]);
    totalDuration += duration;
  } 
  totalDuration = totalDuration/(60000 );
  // Logger.log(totalDuration);
  return totalDuration;
}

function getAvgHexCompleted(data,currentDate,columns){
  for (var i = data.length-1; i > 0 ; i--) {
    // Check if the date in the first column matches the current date
    date = formatDate(new Date(data[i][0]))
    if (date === currentDate && isValid(date)) {
      // Store the last valid row and break out of the loop
      lastValidRow = data[i];
      break;
    }
  }
  // Logger.log(lastValidRow.length);
  weeklyAvgHexCompleted = isValid(lastValidRow[columns.comments+1]) ? lastValidRow[columns.comments+1] : 0;
  // Logger.log(weeklyAvgHexCompleted);
  return weeklyAvgHexCompleted;
}

function includeEmailSignatures(body){
  var emailSignature =
    "<br><br>Regards,<br>" +
    "Dhruval Potla<br>" +
    "RF Engineer<br>" +
    "GCB Services, LLC<br><br>" +
    "Phone     : 703-988-4246/ 703-953-2299 Ext. 107<br>" +
    "Mobile    : (571) 266-0413<br>" +
    "Fax           : 703-738-7676<br><br>" +
    "Delivering solutions for your success.<br><br>" +
    "GCB is among 'Best 5 Telecom companies to watch' for year 2020 & 2023 by Silicon review.<br>" +
    "<a href='https://thesiliconreview.com/magazine/profile/gcb-approach-towards-engineering-and-project-management/'>https://thesiliconreview.com/magazine/profile/gcb-approach-towards-engineering-and-project-management/</a><br>" +
    "<a href='https://thesiliconreview.com/magazine/profile/alter-the-dynamics-of-your-business-with-gcb-service/'>https://thesiliconreview.com/magazine/profile/alter-the-dynamics-of-your-business-with-gcb-service/</a><br><br>" +
    "GCB is among 'Top 10 most promising companies in Wireless Technology' for the year 2019 & 2020 by CIO review.<br>" +
    "<a href='https://wireless.cioreview.com/vendor/2019/gcb_services'>https://wireless.cioreview.com/vendor/2019/gcb_services</a><br>" +
    "<a href='https://wireless.cioreview.com/vendors/top-wireless-technology-consulting-service-companies-2020.html'>https://wireless.cioreview.com/vendors/top-wireless-technology-consulting-service-companies-2020.html</a><br><br>" +
    "GCB is among 'Top 20 most promising companies in Wireless Technology 2017' by CIO review.<br>" +
    "<a href='https://wireless.cioreview.com/vendor/2017/gcb_services'>https://wireless.cioreview.com/vendor/2017/gcb_services</a><br><br>" +
    "<i>This email message is intended by the sender for use only by the individual or entity to which it is addressed. " +
    "This message may contain information that is privileged or confidential. " +
    "It is not intended for transmission to, or receipt by, anyone other than the named addressee (or a person authorized to receive and deliver it to the named addressee). " +
    "If you have received this transmission in error, please delete it from your system without copying or forwarding it, and notify the sender of the error by reply email.</i>";

  body += '\n\n\n';
  body += emailSignature;
  // Logger.log(body);
  return body;
}

function prepareEmailBody(date, hexCompletedToday, idleTSHoursToday, totalIdleTSHoursThisWeek, testersToday, avgHexCompletedThisWeek){
  var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  // date = new Date(date);
  // Find the start date (Monday) by subtracting the current day of the week and adjusting for Monday being the start
  var weekStartDate = new Date(new Date(date).setDate(new Date(date).getDate() - (new Date(date).getDay() + 6) % 7));

  // Find the end date (Sunday) by adding the remaining days until Sunday
  var weekEndDate = new Date(new Date(date).setDate(new Date(date).getDate() + 6));
  var idleTSHoursTodayHours = (idleTSHoursToday/60).toFixed(2) ;

  var html = 'Hi Dish Team,\n\n';
  html += ('Please see below today’s update on GCB activity - ' + date + '\n\n');
  html += '<ul>';
  html += '<li>' + "No. of Hex Visited : " + hexCompletedToday + '</li>';
  html += '<li>' + "Idle/TS hours (" + date + "-" + days[new Date(date).getDay()] + ") : " + idleTSHoursTodayHours + ' hours</li>';
  // html += '<li>' + "Total Idle/TS hours for the week (" + formatDate(weekStartDate) + " to " + formatDate(weekEndDate) + ") : " + totalIdleTSHoursThisWeek + '</li>';
  html += '<li>' + "Total testers in the market : " + testersToday + '</li>';
  html += '<li>' + "Hex completed Average last 7 days : " + avgHexCompletedThisWeek + '</li>';
  html += '</ul>';
  html += '\n'
  html += '<p><span style="font-weight: bold; text-decoration: underline; font-size: larger;">Hex Status:</span></p>';
  html += '\n';
  return html;
}

function sendEmail(to, cc, subject, body, name) {
  // Log the email content
  Logger.log('Email to: ' + to);
  Logger.log('CC: ' + cc);
  Logger.log('Subject: ' + subject);
  Logger.log("Body: " + body);
  // Send email
  MailApp.sendEmail({
    to: to.join(','),
    cc: cc.join(','),
    subject: subject,
    htmlBody: body,
    name : name
  });

  // Log information about the sent email
  Logger.log('Email sent to ' + to + ' with subject: ' + subject);
}



function createHexStatusTable(dataset, tableHtml, headings){
  // Logger.log(dataset.length);
  var hexHeadings = [...headings];
  hexHeadings.push("Duration");
  // Logger.log(hexHeadings);
  // Logger.log(hexHeadings.length);
   
  tableHtml += '<table border="1" style="border-collapse: collapse; width: 100%;"><tr>';
  
  
  for (var i = 0; i < hexHeadings.length; i++) {
    // Logger.log(typeof hexHeadings[i]);
    var value = hexHeadings[i];
    if (i === (hexHeadings.length - 2)){
      value = "Task Status";
    }
    tableHtml += '<th style="background-color: lemon; text-align: center; vertical-align: middle;">' + value + '</th>';
  }
  tableHtml += '</tr>';


  // Loop through the nested array to create table rows
  for (var i = 0; i < dataset.length; i++) {
    tableHtml += '<tr>';
    for (var j = 0; j < hexHeadings.length; j++) {
      var value = '';
      if (j == (hexHeadings.length - 1)){
        startTime = dataset[i][j-3];
        endTime = dataset[i][j-2];
        var durationMs = new Date(endTime - startTime);
        value = Utilities.formatDate(durationMs, "GMT", "HH:mm:ss");
        // Logger.log("Duration: " + value);

      }
      else{
        // Set different background colors for even and odd rows
        value = dataset[i][j];
        // Logger.log(dataset[i][j]);
        if (value instanceof Date) { // If the value is a Date object, format it as a string
          if (hexHeadings[j] === 'Date') {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd/yyyy"); // Format the date as MM/dd/yyyy
          } else if (hexHeadings[j] === 'Start Time' || hexHeadings[j] === 'End Time') {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm"); // Format the time as HH:mm
          }
        }
      }
      
      value = value || '';
      // Logger.log(value);
      var bgColor = i % 2 === 0 ? '#f2f2f2' : '#ffffff';
      tableHtml += '<td style="border: 1px solid #dddddd; font-size: 12px; padding: 4px; background-color: ' + bgColor + ';">' + value + '</td>';
    }
    tableHtml += '</tr>';
  }

  tableHtml += '</table>';
  // Logger.log(tableHtml);
  return tableHtml;
}

function createIdleTSTable(dataset, tableHtml, headings){
  var idTSHeadings = [...headings];
  idTSHeadings.push("Duration");
  // Logger.log(idTSHeadings);

  tableHtml += '\n'
  tableHtml += '<p><span style="font-weight: bold; text-decoration: underline; font-size: larger;">Troubleshoot/Idle Hours:</span></p>';
  tableHtml += '\n';

  tableHtml += '<table border="1" style="border-collapse: collapse; width: 100%;"><tr>';
  
  
  for (var i = 0; i < idTSHeadings.length; i++) {
    var value = idTSHeadings[i];
    if (i === (idTSHeadings.length - 2)){
      value = "Comments";
    }
    tableHtml += '<th style="background-color: lemon; text-align: center; vertical-align: middle;">' + value + '</th>';
  }
  tableHtml += '</tr>';

  // Loop through the nested array to create table rows
  for (var i = 0; i < dataset.length; i++) {
    tableHtml += '<tr>';
    for (var j = 0; j < idTSHeadings.length; j++) {
      var value = '';
      if (j == (idTSHeadings.length - 1)){
        startTime = dataset[i][j-3];
        endTime = dataset[i][j-2];
        var durationMs = new Date(endTime - startTime);
        value = Utilities.formatDate(durationMs, "GMT", "HH:mm:ss");
        // Logger.log("Duration: " + value);

      }
      else{
        // Set different background colors for even and odd rows
        value = dataset[i][j];
        // Logger.log(dataset[i][j]);
        if (value instanceof Date) { // If the value is a Date object, format it as a string
          if (idTSHeadings[j] === 'Date') {
            value = Utilities.formatDate(value, Session.getScriptTimeZone() , "MM/dd/yyyy"); // Format the date as MM/dd/yyyy
          } else if (idTSHeadings[j] === 'Start Time' || idTSHeadings[j] === 'End Time') {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm"); // Format the time as HH:mm
          }
        }
      }

      value = value || '';
      // Logger.log(value);
      var bgColor = i % 2 === 0 ? '#f2f2f2' : '#ffffff';
      tableHtml += '<td style="border: 1px solid #dddddd; font-size: 10px; padding: 4px; background-color: ' + bgColor + ';">' + value + '</td>';

    }
  }
  tableHtml += '</table>';
  return tableHtml;
}

function getIdleTSData(data, columns){
  var dataset = data.filter(function(row) {
      var rowTask = row[columns.taskColumn].toLowerCase();
      // Logger.log(rowTask);
      var result = ((rowTask === 'idle' || rowTask === 'troubleshooting' || rowTask === 'travel_billable') && isValid(rowTask));
      return result;
    });
    // Logger.log(dataset);
    // Logger.log(dataset.length);
    return dataset;
}

function getIdleTSDatahours(data, columns){
  var dataset = data.filter(function(row) {
      var rowTask = row[columns.taskColumn].toLowerCase();
      // Logger.log(rowTask);
      var result = ((rowTask === 'idle' || rowTask === 'troubleshooting') && isValid(rowTask));
      return result;
    });
    // Logger.log(dataset);
    // Logger.log(dataset.length);
    return dataset;
}

function getHexStatusData(data, columns){
  // Logger.log(data);
  var dataset = data.filter(function(row) {
      var rowTask = row[columns.taskColumn].toLowerCase();
      var result = ((rowTask === 'mobility' || rowTask === 'stationary' ) && isValid(rowTask));
      return result;
    });
    return dataset;
}



function getCombinedDateData(data, date, columns) {
  // Logger.log(data);
  var dataset = data.filter(function (row) {
    var rowDate = formatDate(new Date(row[columns.dateColumn]));
    var market = row[columns.marketColumn]
    var tester = row[columns.testerColumn];
    var hex = row[columns.hexNameColumn];
    var task = row[columns.taskColumn];
    var startTime = row[columns.startTimeColumn];
    var endTime = row[columns.endTimeColumn];
    var comments = row[columns.comments];
    return rowDate === date && (isValid(date) && isValid(market) && isValid(tester) && isValid(hex) && isValid(task) && isValid(startTime) && isValid(endTime) && isValid(comments));
  });

  var usefulColumns = Object.values(columns).sort(function(a, b) {return a - b;});
  // Logger.log(usefulColumns);
  var filteredData = dataset.map(function (row) {
    return usefulColumns.map(function (index) {
      return row[index];
    });
  });
  // Logger.log(filteredData);
  return filteredData;
}

// Function to check if a value is valid (modify this based on your criteria)
function isValid(value) {
  return value !== undefined && value !== null && value !== '';
}

function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0');
  var yyyy = date.getFullYear();

  return mm + '/' + dd + '/' + yyyy;
}