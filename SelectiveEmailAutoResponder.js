// Scan new emails and send an autoresponse based on the source address's presence in a spreadsheet.
// The body of the message will also be based on the spreadsheet.

// Requires permissions for Drive/Sheets (Google Sheets API) and EMail (GMail API).
// You need to do this through https://console.cloud.google.com AND at scripts.google.com under Resources/
// Advanced Google Services.

// Last, it's important that you create an EMPTY Google sheet and set the SelectiveEmailSheetURL to that Sheet's
// url.

// Columns: email, responsebody, startdate, enddate
// Data should start on row 2 (so that row 1 can be labels).
SelectiveEMailSheetURL="https://docs.google.com/spreadsheets/d/1zzEPD-jN-HYmGiUwqXtCx_hHW0GN5A299nHoScvman8/edit#gid=0"

function SelectiveEmailAutoResponse(e) {
  var debug = false;
  
  // Use a label to indicate when an autoresponse has been sent.  Create it if it doesn't exist.
  respondedLabelString="autoresponded"
  respondedLabel = GmailApp.getUserLabelByName(respondedLabelString);
  if (!respondedLabel) {
    respondedLabel = GmailApp.createLabel(respondedLabelString);
  }
  
  var ConfigFileName = "AutoResponder Configuration";
  var configFiles = DriveApp.getFilesByName(ConfigFileName);
  
  // If there isn't one, create a sheet.
  if (!configFiles.hasNext()) {
    var ssa = SpreadsheetApp.create(ConfigFileName, 50 , 4);
    
    var header = [["Source Email", "Response Body", "Start Timestamp", "End Timestamp"]];

    ssa.getSheets()[0].getRange("A1:D1").setValues(header);
    
    Logger.log("Created new spreadsheet configuration file : " + ssa.getUrl());
    
    // Stop here and the next run will proceed beyond this step.
    return;
  }
  
  var spreadsheet = SpreadsheetApp.open(configFiles.next());
  var configSheet = spreadsheet.getSheets()[0] 
  
  // If there is another we have a problem (duplicate configs).  Log and bail.
  if (configFiles.hasNext()) {
    Logger.log("ERROR: There are at least two configuration filenames.  Check: " + ssa.getUrl() + " and " + configFiles.next().getUrl());
    // TODO: Stronger way to record an error?
    return;
  }  
  
  var configArray = configSheet.getRange("A2:ZZZ").getValues(); // Get all rows after header in grid (2d array) form.
  
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var dateString = yesterday.getFullYear() + '/' + (yesterday.getMonth()+1) + '/' +yesterday.getDate();
  var accountEmail = Session.getActiveUser().getEmail(); // Used to suppress self auto-respon
 
  // Consider emails only from the last two days.  Using a search filter for this.
  var threads = GmailApp.search("to:me after:" + dateString + " AND !label:" + respondedLabelString);
  
  threadLoop:
  for (var i = 0; i < threads.length; i++) {
    // Parse each candidate email.
    
    var threadTime = threads[i].getLastMessageDate()
    var threadSubject = threads[i].getMessages().reverse()[0].getSubject();
    
    doLog(debug, "RAW TO FIELD: " + threads[i].getMessages().reverse()[0].getTo());
    
    //  getFrom() returns "Some Name <sname@gmail.com>"; regex breaks out the email addr itself.
    var threadFrom = threads[i].getMessages().reverse()[0].getFrom().replace(/.*( |^|<)([^ <]+@[^ >,]+).*>/, "$2");
    var threadTo = threads[i].getMessages().reverse()[0].getTo().replace(/.*( |^|<)([^ <]+@[^ >,]+).*/, "$2");
    doLog(debug, "Processed to field : " + threadTo);
    
    doLog(debug, "Checking email from:" + threadFrom + " to:" + threadTo + " time:" + threadTime + " Subject:" + threadSubject + " RAWTO:" + threads[i].getMessages().reverse()[0].getTo());

    
    // TODO: Support multiple recipients in to field ?
    // Skip emails sent by account owner (unless you want to bounce emails to yourself).
    /* -- In development mode we DO want to reply to ourselves!
    if (accountEmail != threadTo) {
      doLog(debug, "  NIX because " + accountEmail + " == " + threadFrom);
      continue; 
    }
    */
    
    // For each message, check the config sheet for a match.
    configLoop:
    for (var j = 0; j < configArray.length; j++) {
      var configRow = configArray[j];
      
      var emailFrom = configRow[0];
      var responseBody = configRow[1];
      var timeStart = new Date(configRow[2]).getTime();
      var timeEnd = new Date(configRow[3]).getTime();
      
      // Skip rows where the config sender email is blank or config responseBody is blank.
      if (emailFrom == "" || responseBody == "") {
        continue; 
      }
      
      doLog(debug, "  Checking against sender " + emailFrom);

      // Skip messages outside the window for this autoresponse.
      if ((timeStart != "" && threadTime < timeStart) || (timeEnd != "" && threadTime > timeEnd)) {
        doLog(debug, "  NIX.  Time range.");
        continue;
      }
      
      // Skip messages not from the sender we're targetting.
      if (emailFrom != threadFrom) {
        doLog(debug, "  NIX.  EMail mismatch " + emailFrom + " vs " + threadFrom);
        continue; 
      }
      
      doLog(true, "Sending an autoresponse To " + threadFrom + " with body " + responseBody + " on email with subject " + threadSubject)
      threads[i].reply("", {htmlBody: responseBody});
      threads[i].addLabel(respondedLabel);
      
      // Don't send more than one email per thread.
      continue threadLoop;

    }
  }
}

function doLog(doLog, message) {
  if (doLog) {
    Logger.log(message); 
  }
}
