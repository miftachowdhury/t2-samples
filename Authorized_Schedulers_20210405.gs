function onSubmit(e) {
  
 /************************************************************************************************************************************
  *                                                      MAIN EDIT                                                                       *
  ************************************************************************************************************************************/
  
  //open the primary (master) spreadsheet and set directories
  var formUrl = 'https://formfaca.de/sm/f2xAsDbpv';
  var ss = SpreadsheetApp.openById('10PpIsgrIL5tDRr12YqDYU-qahFzViooJJDiNA0zya2Y')
  var dailySheetsDirId = '1b4Zg3cHWi2AuBBeKXw7ra62I983L07nj';                      
  var convertedUploadsDirId = '1XYosjCT9SmjwVUP32thcmwbRcZfSvA0p';
  
  //name primary spreadsheet sheets
  var masterSheet = ss.getSheetByName('Sheet List');
  var validUserSheet = ss.getSheetByName('Valid Users');
  var logSheet = ss.getSheetByName('Processing Log');
  var orgListSheet = ss.getSheetByName('Org List');
  var doittIdSheet = ss.getSheetByName('Orgs from DoITT');
      
  //save respondent information
  var respFirst = e.namedValues['First Name'][0];
  var respLast = e.namedValues['Last Name'][0];
  var respEmail = e.namedValues['Email Address'][0];
  var respOrg = e.namedValues['Name of Organization'][0];
  
  var respInfo = [respFirst, respLast, respEmail, respOrg];
  
  //declare email variables for later use
  var emailMessage = 'N/A';
  var emailTimeSent = new Date();
  var emailSentSuccess = true;
  
  //bulk upload variables for later use
  var dataFileName = '';
  var dataFileUrl = '';
    
  //individual add variables for later use
  var [userFirst, userLast, userEmail, userOrg] = ['', '', '', ''];
  var [origBatchId, origRespFirst, origRespLast, origTime] = ['', '', '', null];  
  
  //format timestamp for readability
  var timeStamp = new Date(e.namedValues.Timestamp[0]);
  var lastSubmitTime = Utilities.formatDate(timeStamp, ss.getSpreadsheetTimeZone(), "M/d/yy h:mm a");
  var emailDate = Utilities.formatDate(timeStamp, ss.getSpreadsheetTimeZone(), "M/d/yy 'at' h:mm a");
  
  //get the last unique ID and set the next one; if no data yet, then set the next one to 'N-00001'
  var UID = ''
  
  if (validUserSheet.getLastRow() == 1) {
    var UID = 'N-00001';
  } else if (validUserSheet.getLastRow() > 1) {
    var lastUIDstring = validUserSheet.getRange(validUserSheet.getLastRow(), 3).getValue();
    var UID = 'N-' + (parseInt(lastUIDstring.slice(2)) + 1).toString().padStart(5, '0');
  }
  
  //check the date column of the primary spreadsheet for today's date
  var date = new Date();
  var minutes = date.getMinutes() + 300;
  date.setMinutes(minutes);
  var newDate = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd").toString();
  
  var masterRows = masterSheet.getLastRow()+1;
  
  for (var row = 1; row < masterRows; row++) {  
    if (masterSheet.getRange(row,1).getDisplayValue() == newDate) {
      var sheetExists = true;
      var sheetRow = row;
      break;
    } 
    else {
      var sheetExists = false;
      var sheetRow = masterRows;
    }
  }

  Logger.log('The sheet for today exists in the Date column: ' + sheetExists);
     
  //if today's spreadsheet does not exist then create it
  if (sheetExists == false) {
    createSpreadsheet(newDate);    
  }
  
  //identify any cells with values that will need to be used or updated
  var urlCell = masterSheet.getRange(sheetRow, 3);
  var lastSubmitCell = masterSheet.getRange(sheetRow, 4);
  var numSubmitCell = masterSheet.getRange(sheetRow, 5);
  var numUsersCell = masterSheet.getRange(sheetRow, 6);
  
  //store existing values and declare variables
  var numSubmit = numSubmitCell.getValue();
  var numUsers = numUsersCell.getValue();
  var rawData = [];
  var batchId = '';
  var submitUsers = null;
  
  //check if the submission is bulk or individual, then invoke the appropriate function to get user account data
  var bulk = e.namedValues['How many accounts do you need created?'] == 'I need to create more than 1 new user account';
  var selfRequest = e.namedValues['How many accounts do you need created?'] == 'I need one account created for myself';
  var noAccount = e.namedValues['How many accounts do you need created?'] == 'I do not need to create any accounts at this time.'; 
  var isIndivDupe = false;
  
  Logger.log('This is a bulk submission: ' + bulk);
  Logger.log('This is a self request: ' + selfRequest);
  
  if (noAccount) {
    Logger.log('User indicated no accounts are needed.');
    return;
  }
  
  if (bulk) {
    
    bulkAdd();
    var dataToValidate = preProcessData(rawData);
    var [isValid, validResults]  = validateData(dataToValidate);
     
    if (isValid) {
      writeData();
    } 
    
  } else {   
    
    indivAdd();
    
   if (!isIndivDupe) {
      writeData(); 
    }
  }
  
  tryEmail();
  
  /**********************************************************************************************************************************/
  /**********************************************************************************************************************************/
  
  //save processing information to the processing log
  var metaData = [lastSubmitTime, newDate, batchId]
  var emailInfo = [emailMessage, emailSentSuccess, emailTimeSent]
  
  saveToLog(metaData, respInfo, validResults, emailInfo);
  
  function saveToLog(metaData, respInfo, validObj, emailInfo) {
    var logArray = [];
    logArray = logArray.concat(metaData, respInfo);
    if (!bulk) {
      logArray.push('N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A');
    } else {
      logArray.push(dataFileUrl); //bulk upload file url
      logArray = logArray.concat(Object.values(validObj).map(x => x.valid)); //missing, email, and org name validation results
      logArray.push(Object.values(validObj).map(x => x.valid).reduce((x,y) => x && y)); //overall validation result
      logArray.push(validObj); //validation object
    }
    logArray = logArray.concat(emailInfo); //email related info
    logSheet.appendRow(logArray);
}
  
  //write the data to the daily sheet as well as the all users sheet and update the sheet list metadata
  function writeData() {
    //update values in the primary sheet and create a submission ID
    numUsers += submitUsers;
    numSubmit++;  
    
    lastSubmitCell.setValue(Utilities.formatDate(timeStamp, ss.getSpreadsheetTimeZone(), "M/d/yy h:mm a"));
    numSubmitCell.setValue(numSubmit);
    numUsersCell.setValue(numUsers);
    batchId = createBatchId(newDate, numSubmit);   
    
    //open the spreadsheet
    var getUrl = urlCell.getValue();
    var todaySS = SpreadsheetApp.openByUrl(getUrl);
    var formSheet = todaySS.getSheetByName('Users');
    
    for (i = 1; i < submitUsers+1; i++) {
      
      //write to daily sheet
      var dataArray = rawData[i];
      var org = dataArray[3];
      var orgId = getOrgId(org);
      dataArray.push(orgId);
      formSheet.appendRow(dataArray);
      
      //write to valid users sheet
      var validUserArray = dataArray;
      validUserArray.unshift(UID);
      validUserArray.unshift(batchId);
      validUserArray.unshift(newDate);
      validUserSheet.appendRow(validUserArray);
      
      UID = 'N-' + (parseInt(UID.slice(2)) + 1).toString().padStart(5, '0');
            
    }
  }

  //look up and get the organization's enroller ID
  function getOrgId (userOrg) {
    var orgRows = doittIdSheet.getLastRow()+1;
    for (var row = 1; row < orgRows; row++) {  
      if (doittIdSheet.getRange(row,2).getValue() == userOrg) {
        var orgId = doittIdSheet.getRange(row, 4).getValue();
        break;
      } 
    }
    Logger.log('Organization ID: ' + orgId);
    return orgId;
  }  
  
  
 /************************************************************************************************************************************
  *                                              BULK UPLOAD FUNCTION                                                                *
  ************************************************************************************************************************************/
  
  //function to invoke if user selects bulk add option
  function bulkAdd() {
    
    var ssId;
    var folderId = convertedUploadsDirId;
    //get the datafile ID and then the datafile
    dataFileUrl = e.namedValues['Please upload spreadsheet here.'][0];
    
    Logger.log('File URL: ' + dataFileUrl);
    
    var response = UrlFetchApp.fetch(dataFileUrl);
    var excelFile = DriveApp.createFile(response);
    dataFileName = excelFile.getName();
    Logger.log('Data File Name 1: ' + dataFileName);
    //dataFileName = dataFileName.replace(/%20/g, ' ');
    dataFileName = decodeURIComponent(dataFileName);
    Logger.log('Data File Name 2: ' + dataFileName);
    var blob = excelFile.getBlob();
    
    var driveresource = 
        {
          title: dataFileName,
          mimeType: MimeType.GOOGLE_SHEETS,
          parents: [{id: folderId}]
        };
    
    var newFile = Drive.Files.insert(driveresource, blob);
    ssId = newFile.id;
    Logger.log('ssId: ' + ssId);
    
    //open the file and get the data
    var spreadsheet = SpreadsheetApp.openById(ssId);
    var fileSheet = spreadsheet.getSheets()[0];
    rawData = fileSheet.getDataRange().getValues();
            
    //store the number of users in the submission
    submitUsers = rawData.length-1
  }
  
  function preProcessData(data) {
    var dataDict = {};
    for (var colIndex in data[0]) {
      var colName = data[0][colIndex];
      var colVals = data.slice(1).map(val => val[colIndex]);
      dataDict[colName] = colVals;
    }
    return dataDict;
  }
  

 /************************************************************************************************************************************
  *                                            INDIVIDUAL ACCOUNT FUNCTION                                                           *
  ************************************************************************************************************************************/

  //function to invoke if user selects individual sign-up option  
  function indivAdd() {
    
    if (selfRequest) {
      userFirst = respFirst;
      userLast = respLast;
      userEmail = respEmail;
      userOrg = respOrg;
    } else {
      userFirst = e.namedValues['User First Name'][0];
      userLast = e.namedValues['User Last Name'][0];
      userEmail = e.namedValues['User Email Address'][0];
      userOrg = e.namedValues['Organization Name'][0];
    }
    
    //check if the submission is a duplicate
    checkIndivDupe();
    
    //if the submission is not a duplicate, then...
    if (!isIndivDupe) {
      //save user information in two-dimensional array
      rawData = [ [], [userFirst, userLast, userEmail, userOrg]];
      //and store the number of users in the submission
      submitUsers = 1;
    }

  }
 /************************************************************************************************************************************/  
  
  /**** Create batch ID ****/
  function createBatchId (date, num) {
    function pad(n) {
      return (n < 10) ? ('0' + n) : n.toString();
    }
    var stringNum = pad(num);
    var string = 'B-' + date.slice(4).replace(/-/g, '') + '-' + stringNum; 
    return string;
  }
  
  
 /************************************************************************************************************************************
  *                                             CREATE THE DAY'S SPREADSHEET                                                         *
  ************************************************************************************************************************************/
  function createSpreadsheet(sheetDate) {
    
    //create today's spreadsheet and set column headers
    var todayFileName = 'Authorized_Scheduler_Accounts-New_Users_' + sheetDate;
    var todaySS = SpreadsheetApp.create(todayFileName);
    var ssFile = DriveApp.getFilesByName(todayFileName).next();
    var dailySheetsDir = DriveApp.getFolderById(dailySheetsDirId)
    ssFile.moveTo(dailySheetsDir);
    todaySS.renameActiveSheet('Users');
    var formSheet = todaySS.getSheetByName('Users');
    formSheet.appendRow(['First Name', 'Last Name', 'Email', 'Organization Name', 'CBO/Enroller ID']);
    
    //get the urls for csv and xlsx formats of the day's spreadsheet
    var csvUrl = 'https://docs.google.com/spreadsheets/d/' + todaySS.getId() + '/export?format=csv&gid=' + formSheet.getSheetId();
    var xlsxUrl = 'https://docs.google.com/spreadsheets/d/' + todaySS.getId() + '/export?format=xlsx';
    
    //and write into the primary sheet: the date, file name, spreadsheet ID, spreadsheet URL, CSV URL, XLSX URL, and additional info
      masterSheet.appendRow([sheetDate, 
                             todayFileName, 
                             todaySS.getUrl(),
                             lastSubmitTime,
                             null,
                             null,
                             csvUrl, 
                             xlsxUrl]);
  }

  
 /************************************************************************************************************************************
  *                                            VALIDATION FUNCTIONS                                                                  *
  ************************************************************************************************************************************/  
  
  //validate overall 
  function validateData(dataDict) {
    var validDict = {
    'valid_required_obj': validateRequired(dataDict, ['First Name', 'Last Name', 'Email', 'Organization Name']), 
    'valid_email_obj': validateEmail(dataDict, 'Email', 'First Name', 'Last Name'),
    'valid_org_obj': validateOrgName(dataDict, 'Organization Name')
    };
    
    // reduce to get overall valid
    var isValidAll = Object.values(validDict).map(x => x.valid).reduce((x,y) => x && y); 
    return [isValidAll, validDict];
  }
  
  
  //validate columns not missing data 
  function validateRequired(data, requiredCols) {
    var isValid = true;
    var failString = '<br>The following required columns are missing values:<br>'
    for (var reqCol of requiredCols) {
      if (missingColumnData(data[reqCol])) {
        isValid = false;
        failString += '   -' + reqCol + '<br>';
      } else if (data[reqCol].length != data[reqCol].filter(x => x != '' && x != null).length) {
        isValid = false;
        failString += '   -' + reqCol + '<br>';
      }
    }
    failString += '<br>Please check your spreadsheet and make sure the columns match the template exactly.<br>'
    return { 'valid': isValid, 'failString': failString };
  }
  
  //validate email address
  function validateEmail(data, emailColName, firstColName, lastColName) {
    var emailColArr = data[emailColName];
    var firstColArr = data[firstColName];
    var lastColArr = data[lastColName];
    
    if (missingColumnData(emailColArr)) return { valid: true };
    var isValid = true;
    var failNames = [];
    var failString = '<br>The email addresses provided for the following individuals are invalid:<br>';
    
    //check that email address is in a valid email address format
    emailColArr.forEach(checkForValidEmail);
    
    function checkForValidEmail(email, index) {
      var validEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
      if(!validEmail) {
        isValid = false;
        var fullName = firstColArr[index] + ' ' + lastColArr[index];
        failNames.push(fullName);
        failString += '   -' + fullName + '<br>';
      }
    }
    
    failString += '<br>Please check that all email addresses are valid.<br>';
    return { 'valid': isValid, 'failNames': failNames, 'failString': failString };
  }
  
  //validate organization name
  function validateOrgName (data, orgColName) { 
    var orgColArr = data[orgColName];
   
    if (missingColumnData(orgColArr)) return { valid: true };
    var isValid = true;
    var failOrgs = [];
    var failString = '<br>The following organization names provided are not included in the list of participating CBOs:<br>';
    
    //for each unique org name in the data, check that the org name is found in the org list
    var uniqueOrgs = [...new Set(orgColArr)];
    uniqueOrgs.forEach(checkForOrgName);
    
    function checkForOrgName (orgName) {
      //get the list of participating CBOs
      var definedOrgArray = orgListSheet.getDataRange().getValues();
      definedOrgArray = definedOrgArray.map(org => org[1]);

      var validOrg = definedOrgArray.includes(orgName);
      if (!validOrg) {
        isValid = false;
        failOrgs.push(orgName);
        failString += '   -' + orgName + '<br>';
      }
    }
    
    failString += '<br>Please check that all organization names are spelled correctly and included in the list of participating CBOs.<br>';
    return { 'valid': isValid, 'failOrgs': failOrgs, 'failString': failString };
  }
  
  function missingColumnData(columnArr) { return (columnArr == null || columnArr.length == 0) }
  
  
 /************************************************************************************************************************************
  *                                            DUPLICATE CHECK FUNCTIONS                                                             *
  ************************************************************************************************************************************/   
  
  //function to check if an individual submission is a duplicate
    function checkIndivDupe() {
      var firstInstance = true;
      var dupeCount = 0;
          
      var validUserRows = validUserSheet.getLastRow() + 1;
      for (var row = 1; row < validUserRows; row++) {
        if (validUserSheet.getRange(row, 6).getValue() == userEmail) {
          if (firstInstance) {
            isIndivDupe = true;
            //get original batch ID
            origBatchId = validUserSheet.getRange(row, 2).getValue();
            origUID = validUserSheet.getRange(row, 3).getValue();
            origUserFirst = validUserSheet.getRange(row, 4).getValue();
            origUserLast = validUserSheet.getRange(row, 5).getValue();
            origUserEmail = validUserSheet.getRange(row, 6).getValue();
            origUserOrg = validUserSheet.getRange(row, 7).getValue();
            firstInstance = false;
          }
          //increment duplicate count
          dupeCount++;
        }
      }
      
      Logger.log('This a duplicate 1: ' + isIndivDupe);
      Logger.log('The original Batch ID is: ' + origBatchId);
      
      if (isIndivDupe) {
        //get info about original submission using original batch ID
        
        var logSheetRows = logSheet.getLastRow() + 1;
        for (var logRow = 1; logRow < logSheetRows; logRow++) {
          if (logSheet.getRange(logRow, 3).getValue() == origBatchId) {
            origRespFirst = logSheet.getRange(logRow, 4).getValue();
            origRespLast = logSheet.getRange(logRow, 5).getValue();
            origRespName = origRespFirst + ' ' + origRespLast;
            origRespEmail = logSheet.getRange(logRow, 6).getValue();
            origTime = logSheet.getRange(logRow, 1).getValue();
            origTime = Utilities.formatDate(origTime, ss.getSpreadsheetTimeZone(), "M/d/yy 'at' h:mm a");
          }
        }
                            
      }
      
    } 
    
  
  
 /************************************************************************************************************************************
  *                                                EMAIL FUNCTIONS                                                                   *
  ************************************************************************************************************************************/  
   
  //create and send a response (confirmation/notification) email
  function sendResponseEmail (recipientEmail, errorObj, dataFileName, submissionTime, formUrl) {
    Logger.log('Function sendResponseEmail has been reached.');
    var successSubject = 'Authorized Vaccination Scheduler Form - Successful Submission';
    var failSubject = 'Authorized Vaccination Scheduler Form - Submission Error';
    var dupeSubject = 'Authorized Vaccination Scheduler Form - Duplicate Request Received';
    var subject = '';
    var message = '';
        
    //if submission was for an individual user account...
    if (!bulk) {
    Logger.log('Not bulk.');
      //...and submission was not a duplicate, then compose success email
      if (!isIndivDupe) {      
        subject = successSubject;
        var userInfo =  'User Name: ' + userFirst + ' ' + userLast + '<br>User Email: ' + userEmail + '<br>User Organization: ' + userOrg;
        message = "Your submission on " + submissionTime + " to the " + '<a href=\"' + formUrl + '">Authorized Schedulers New User Signup</a>' + " form was successful!<br><br>The following information has been received:<br><br>" + userInfo + "<br><br>Please note: It may take up to 24 hours for users to receive their account log-in credentials. Credentials will be sent directly to the email address(es) you provided. Please do not resubmit this form.<br><br>If after 24 hours user(s) have not received the account credentials, first, users should check their spam folders. If nothing is there, please reply and let us know the accounts that are missing.<br><br>-Vaccine Command Center Equity Team<br>"; 
      
      //...and submission was a duplicate, then compose duplicate email 
      } else {
        Logger.log('Is dupe');
        subject = dupeSubject;
        message = "An account has already been requested for the email address " + userEmail + ". The original request was made by " + origRespFirst + " " + origRespLast + " on " + origTime + " with the following information:<br><br>"
        + "User Name: " + origUserFirst + " " + origUserLast + "<br>User Email: " + origUserEmail + "<br>User Organization: " + origUserOrg + "<br><br>"
        + "Please note: It may take up to 24 hours for users to receive their account log-in credentials. Credentials will be sent directly to the email address provided. Please do not resubmit this form.<br><br>If after 24 hours user(s) have not received the account credentials, first, users should check their spam folders. If nothing is there, please reply and let us know the accounts that are missing.<br><br>-Vaccine Command Center Equity Team<br>";
        Logger.log('Message: ' + message);
      }    
      
    //if submission was a bulk upload...
    } else if (bulk) {
      
      //...and submission was valid, then compose success email
      if (isValid) {
        subject = successSubject;
        message = "Your submission of " + dataFileName + " on " + submissionTime + " to the " + '<a href=\"' + formUrl + '">Authorized Schedulers New User Signup</a>' + " form was successful!<br><br>Please note: It may take up to 24 hours for users to receive their account log-in credentials. Credentials will be sent directly to the email address(es) you provided. Please do not resubmit this form.<br><br>If after 24 hours user(s) have not received the account credentials, first, users should check their spam folders. If nothing is there, please reply and let us know the accounts that are missing.<br><br>-Vaccine Command Center Equity Team<br>";
      
      //...and submission was invalid, then compose fail email  
      } else {
        
        subject = failSubject;
        
        // combine failStrings into one message
        var errorString = '';
        Object.values(errorObj).forEach(validObj => {
          if (!validObj.valid) {
          errorString += validObj.failString;
          }
        });
        
        //set the message
        message = "Your submission of " + dataFileName + " on " + submissionTime + " to the Authorized Schedulers New User Signup form failed for the following reasons:<br>" + errorString + "<br>Please resubmit with corrections at " + formUrl + ", thank you!<br><br>-Vaccine Command Center Equity Team<br>";          
      } //close invalid                           
           
    } //close bulk
    
    //set the components of the email
    var emailTemplate = {
      to: recipientEmail,
      subject: subject,
      htmlBody: message,
      name: 'NYC Vaccine Command Center Equity Team',
      replyTo: 'VaccineEquity@cityhall.nyc.gov'
    }
    
    //send the email
    MailApp.sendEmail(emailTemplate);
    return message;  
  }

  //try to send the email
  function tryEmail() {
    emailTimeSent = new Date();
    emailSentSuccess = true;
    try {
      emailMessage = sendResponseEmail(respEmail, validResults, dataFileName, emailDate, formUrl);
    } catch (error) {
      emailSentSuccess = false;
    }
  }
      
   
}
