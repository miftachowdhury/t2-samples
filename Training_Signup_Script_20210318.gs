function onTrainingSubmit(e) {

  //read in user-submitted values from form submission
  var training = e.namedValues['Scheduled Trainings'][0];
  var firstName = e.namedValues['First Name'][0];
  var lastName = e.namedValues['Last Name'][0];
  var email = e.namedValues['Email'][0];
  var org = e.namedValues['Organization'][0];
  var fullName = firstName + ' ' + lastName;
  
  //declare spreadsheet variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  //declare email variables
  var emailMessage = '';
  var emailTimeSent = null;
  var emailSentSuccess = true;
  
  //format timestamp
  var timeStamp = new Date(e.namedValues.Timestamp[0]);
  var signUpTime = Utilities.formatDate(timeStamp, ss.getSpreadsheetTimeZone(), "M/d/yy h:mm a");
  
  //get training date as a date object
  var trainingDateObj = new Date(training);
    
  //parse the training name as a date string
  var dateSplitArray = training.split(" ");
  var [day, date, time, amPm] = dateSplitArray;
  var fullWeekday = Utilities.formatDate(trainingDateObj, ss.getSpreadsheetTimeZone(), "EEEE");
  
  //distinguish timeslot sheets from non timeslot sheets
  var nonSlotSheets = ['Form Responses 1', 'Scheduled Trainings', 'Processing Log', 'Org List'];
  var slotSheets = [];
  
  sheets.forEach(function(sheet){
    if (!nonSlotSheets.includes(sheet.getSheetName())) {
      slotSheets.push(sheet);
    }
  });
  
  //create an array of all the timeslot sheet names
  var slotSheetNames = [];
  slotSheets.forEach(function(sheet) {
    slotSheetNames.push(sheet.getSheetName());
  });
  
  //if there are no timeslot sheets, create the sheet and insert it as the first sheet, then add the row
  if (slotSheets.length == 0) {
    createNewSlotSheet(ss, training);
    addEntry(ss, training);
    
  //if there is at least one timeslot sheet then
  } else if (slotSheets.length > 0) {
       
    //if the list of timeslot sheets does not already contain the sheet for this training, then create it and add the row
    if (!slotSheetNames.includes(training)) {
      createNewSlotSheet(ss, training);
      sortSlotSheets(ss, training);
      addEntry(ss, training);
      
      //otherwise, add the row
    } else {
      addEntry(ss, training);
    }
    
  }  
  
  //try to send a confirmation email
  tryEmail(email, time);
  
  //save results to processing log
  saveToLog(ss);     
    
    
/*****************************************************************************************************************************/
  function sendConfirmEmail(email, time) {
    
    //get the correct link to the training
    var scheduleSheet = ss.getSheetByName('Scheduled Trainings');
    var timeSlotFinder = scheduleSheet.createTextFinder('Time Slot');
    var timeSlotCell = timeSlotFinder.findNext().getCell(1,1);
    var timeSlotRng = timeSlotCell.getDataRegion();
    
    //get the training link, phone number, and meeting ID
    if (time == '10:00') {
      var url = timeSlotRng.getCell(2,2).getValue();
      var phoneNum = timeSlotRng.getCell(2,3).getValue();
      var meetingId = timeSlotRng.getCell(2,4).getValue();
    } else if (time == '3:00') {
      var url = timeSlotRng.getCell(3, 2).getValue();
      var phoneNum = timeSlotRng.getCell(3, 3).getValue();
      var meetingId = timeSlotRng.getCell(3, 4).getValue();
    }    
    
    //compose the email
    var subject = "Confirmation: Upcoming Authorized Scheduler Training on " + fullWeekday + " " + date + " at " + time + " " + amPm;
    
    var body = "If you're receiving this email, we are confirming your attendance at the upcoming Authorized Scheduler Training on "
    + fullWeekday + " " + date + " at " + time + " " + amPm + "."
    + "<br><br>"
    + "In order to join the training, please go to the following link: " + url + "."
    + "You may also join by phone, by dialing " + phoneNum + " and entering the following meeting ID: " + meetingId + "."
    + "<br><br>"
    + "We will be presenting slides, so we strongly encourage you to join via computer if possible. "
    + "You will receive a copy of the training deck, along with an FAQ and other training materials, following the presentation"
    + "<br><br>" + "Thank you," + "<br>" + "Vaccine Command Center Equity Team";
    
    var messageAsPlainText = body.replace(/<br>/g, '\n');
    
    var message = {
      to: email,
      subject: subject,
      htmlBody: body,
      name: 'NYC Vaccine Command Center Equity Team',
      replyTo: 'VaccineEquity@cityhall.nyc.gov'
    }
    
    //send the email
    //MailApp.sendEmail(message);
    return messageAsPlainText;  
  }
    
    
/*****************************************************************************************************************************/
  function saveToLog(ss) {
    
    //get the processing log sheet
    var logSheet = ss.getSheetByName('Processing Log');
    
    //create a unique ID
    var UID = ''
    if (logSheet.getLastRow() == 1) {
      var UID = 'N-00001';
    } else if (logSheet.getLastRow() > 1) {
      var lastUIDstring = logSheet.getRange(logSheet.getLastRow(), 1).getValue();
      var UID = 'N-' + (parseInt(lastUIDstring.slice(2)) + 1).toString().padStart(5, '0');
    }
    
    var logRow = [UID, signUpTime, training, firstName, lastName, fullName, email, org, emailMessage, emailSentSuccess, emailTimeSent];
    logSheet.appendRow(logRow);   
  
  }
    
    
/*****************************************************************************************************************************/
  function addEntry(ss, training) {
    var trainingSheet = ss.getSheetByName(training);     
    
    //string together email addresses
    var emailFinder = trainingSheet.createTextFinder('Email');
    var emailCol = emailFinder.findNext().getColumn();
    var emailVals = trainingSheet.getRange(1, emailCol).getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
    var emailArray = [];
    var emailString = '';
    
    if (emailVals.length == 1) {
      emailString = email;
    } else if (emailVals.length == 2) {
      emailString = emailVals[1][0] + '; ' + email;
    } else if (emailVals.length > 2) {
      for (i = 1; i < emailVals.length; i++) {
        emailArray.push(emailVals[i][0]);
      }
      emailString = emailArray.join('; ') + '; ' + email;
    }
    
    //compose a subject line and create the mailto link
    var subject = 'Information%20Regarding%20Authorized%20User%20Training%20on%20' + fullWeekday + '%20' + date + '%20at%20' + time + '%20' + amPm; 
    var emailLink = '=HYPERLINK("mailto: place@holder.com?bcc=' + emailString + '&' + 'subject=' + subject + '", "Send Email to ' + training + ' Training Attendees")';
    
    //write the row data
    var row = [training, signUpTime, firstName, lastName, fullName, email, org];
    trainingSheet.appendRow(row);
    
    //write or update the email link
    var linkFinder = trainingSheet.createTextFinder('Link to Compose Email');
    var linkCol = linkFinder.findNext().getColumn();
    var linkCell = trainingSheet.getRange(2, linkCol);
    linkCell.setValue(emailLink);
  }
    
    
/*****************************************************************************************************************************/
  function createNewSlotSheet(ss, training) {
    //insert the sheet
    ss.insertSheet(training, 0);
    var trainingSheet = ss.getSheetByName(training);
    var headerRng = trainingSheet.getRange("A1:H1");
    var headerVals = [['Training Slot', 'Sign-up Time', 'First Name', 'Last Name', 'Full Name', 'Email', 'Organization', 'Link to Compose Email']];
    headerRng.setValues(headerVals);
  }
  
/*****************************************************************************************************************************/
  function sortSlotSheets(ss, training) {
    var dateTimeExtracts = [];
    var trainingSheet = ss.getSheetByName(training);
    var trainingDate = training.slice(4);
    var trainingAsDate = Date.parse(trainingDate);
    
    //get all the timeslot sheet names as date objects
    slotSheets.forEach(function(sheet){
      var sheetName = sheet.getSheetName();
      var trainingTime = Date.parse(sheetName);
      var dateTimeObj = {name: sheetName, datetime: trainingTime, index: sheet.getIndex()};
      dateTimeExtracts.push(dateTimeObj);        
    });
    
    //var currentDateTimeObj = {name: training, datetime: trainingAsDate, index: trainingSheet.getIndex()};
    //dateTimeExtracts.push(currentDateTimeObj);
    
    //sort the array of date objects
    Logger.log(dateTimeExtracts);
    dateTimeExtracts.sort((a, b) => a.datetime-b.datetime);
    Logger.log(dateTimeExtracts);
    
    //set the training sheet's index
    for (i = 0; i < dateTimeExtracts.length; i++) {
      var obj = dateTimeExtracts[i];
      if (trainingAsDate < obj.datetime) {
        var trainingIndex = obj.index-1;
        break;
      }
      var trainingIndex = dateTimeExtracts.length + 1;
    }
    ss.setActiveSheet(trainingSheet);
    ss.moveActiveSheet(trainingIndex);
  }
  
/*****************************************************************************************************************************/
  
  //try to send the email
  function tryEmail(email, time) {
    emailTimeSent = new Date();
    emailSentSuccess = true;
    try {
      emailMessage = sendConfirmEmail(email, time);
    } catch (error) {
      emailSentSuccess = false;
    }
  }
  
  
}
  

