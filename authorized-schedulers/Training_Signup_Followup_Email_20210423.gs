function sendFollowUpEmail() {
  
  var ss = SpreadsheetApp.openById('1ZoXQYkObFegexv52svKVVWlftKGjX0qn8dpZaHIjZew');
  var todayDate = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "EEE M/d/yy");
  var logSheet = ss.getSheetByName('Follow-up Email Log');
  todayDate = todayDate.toString();
  Logger.log('todayDate: ' + todayDate);

  
  //var dayArray = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
  //var runDays = ['Monday', 'Tuesday', 'Friday']
  //var todayWeekday = dayArray[todayDate.getDay()];
  
  //if (!runDays.includes(todayWeekday)) {
  //  Logger.log('Today is ' + todayWeekday + '. There are no trainings scheduled on ' + todayWeekday + 's. No action taken.');
  //  return;
  //}

  var allSheets = ss.getSheets();
  var todaySheets = []
  
  allSheets.forEach(function(sheet) {
    var sheetName = sheet.getSheetName();
    if (sheetName.includes(todayDate)) {
      todaySheets.push(sheet)
    }
  });
  
  
  if (todaySheets.length == 0) {
    Logger.log('There were no trainings scheduled for today. No action taken.');
    return;
  }
  
  var emailArray = ['mchowdhury1@health.nyc.gov']
  
  todaySheets.forEach(function(tdSheet) {
    var emailFinder = tdSheet.createTextFinder('Email');
    var emailCol = emailFinder.findNext().getColumn();
    var emailVals = tdSheet.getRange(1, emailCol).getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
    for (i = 1; i < emailVals.length; i++) {
        emailArray.push(emailVals[i][0]);
      }
  });

  emailArray = [...new Set(emailArray)];
  var recipients = emailArray.join(',');
  Logger.log('recipients: ' + recipients);
  
  var resourceColFinder = logSheet.createTextFinder('Resource');
  var resourceCol = resourceColFinder.findNext().getColumn();
  var urlCol = resourceCol + 1;
    
  var videoUrl = logSheet.getRange(2, urlCol).getValue();
  var presUrl = logSheet.getRange(3, urlCol).getValue();
  var scriptUrl = logSheet.getRange(4, urlCol).getValue();
  var faqUrl = logSheet.getRange(5, urlCol).getValue();
  var rolesUrl = logSheet.getRange(6, urlCol).getValue();

  //compose and send email
  
  var subject = 'Meeting Follow-up: Vaccine Authorized Scheduler Training';
  
  var body = 'Dear Partners,'
    + '<br><br>'
    + 'Thank you for attending the Vaccine Authorized Scheduler Training!'
    + '<br><br>'
    + '<i>Note, if you have received this email in error and have not yet attended a training, you may sign up for an upcoming training through the link '   
    + '<a href = "https://docs.google.com/forms/d/e/1FAIpQLSfOs_Fu6FF_QL5Gs0-6djqAJPYgMZEum76w1kM5Ev1W6x_PWg/viewform">here</a></i>.'
    + '<br><br>'
    + 'As a reminder, you will receive an email from the NYC Department of Information Technology and Telecommunications (DoITT) with log-in credentials by <b><u>tomorrow</b></u> 6:00pm.'
    + ' If you do not receive an email, first check your spam and/or junk folders.'
    + '<br><br>'
    + 'If there is nothing there, request a log-in through this form: '
    + '<a href = "https://formfaca.de/sm/f2xAsDbpv">https://formfaca.de/sm/f2xAsDbpv</a>.'
    + ' If your organization is not yet showing up on this list, we are awaiting a signed agreement and cannot create your accounts until then.'
    + ' Please continue checking back until your organization is added. Note, your organization may be listed under an umbrella organization.'
    + '<br><br>'
    + 'Upon receiving their credentials, new users should:'
    + '<br><ul>'
    + '<li>Attempt to log-in to your account by entering your username and password. If you have a problem with your password, please use the form on the login screen to contact the support group directly for password issues</li>'
    + '<li>Reset your password</li>'
    + '<li>Confirm that you are assigned to the correct organization</li>'
    + '<li>If you have been assigned to the wrong organization, please email <a href = "mailto:leedsae@nychhc.org">leedsae@nychhc.org</a>.</li>'
    + '</ul><br>'
    + '<b>Training Support</b>'
    + '<br><br>'
    + 'As a follow up to today&#39;s training, please find a list of resources here:'
    + '<br><ul>'
    + '<li>Watch the video of the training, <a href = "' + videoUrl + '">linked here</a>.</li>'
    + '<li>Review today&#39;s training presentation, linked <a href = "' + presUrl + '">here</a>.</li>'
    + '<li>Review the suggested script to help facilitate outreach to your clients. '
    + 'It follows the order of the NYC VAX app, so should simply help you move through screens in the process. '
    + 'Suggested script linked <a href = "' + scriptUrl + '">here</a>.</li>' 
    + '<li>Review the FAQ to help answer common questions, linked <a href = "' + faqUrl + '">here</a>.</li>'
    + '<li>Review the Roles and Responsibilities of Authorized Schedulers, linked <a href = "' + rolesUrl + '">here</a>.</li>'
    + '</ul><br>'
    + 'Please do not hesitate to reach out with any questions! We are so glad to have you on board as a partner.'
    + '<br><br>'
    + 'For questions related to the training or app, please email VaccineEquity@cityhall.nyc.gov.'
    + '<br><br>'
    + 'Thank you for your continued support in protecting and promoting the health of all New Yorkers.'
    + '<br><br>' + 'All the best,' + '<br>' + 'Vaccine Command Center Equity Team';
     
    var messageAsPlainText = body.replace(/<br>/g, '\n'); 

    var message = {
      to: 'VaccineEquity@cityhall.nyc.gov',
      bcc: recipients,
      subject: subject,
      htmlBody: body,
      name: 'NYC Vaccine Command Center Equity Team',
      replyTo: 'VaccineEquity@cityhall.nyc.gov'
    }
    
    Logger.log(messageAsPlainText);
    
    //send the email
    MailApp.sendEmail(message);

    //log the info
    var logTime = new Date();
    logTime = Utilities.formatDate(logTime, ss.getSpreadsheetTimeZone(), "M/d/yy h:mm a");
    
    logSheet.appendRow([todayDate, recipients, logTime]);
  
  
}
