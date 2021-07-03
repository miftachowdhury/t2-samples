//@NotOnlyCurrentDoc

//Purpose:
// Create event on borough-specific calendar and send invitations to staff attending
// Send email notifications conditionally, based on resource needs 

//Create a function to be triggered on form submit
function onFormSubmit(e) {
  
  var form = FormApp.getActiveForm();
  
  //create form response variables
  var formResponse = e.response;
  var itemResponses = formResponse.getItemResponses();
  var respUrl = formResponse.getEditResponseUrl();
  var respID = formResponse.getId();
  form.setAllowResponseEdits(true);
  
  //create spreadsheet variables - workbook and sheet
  var ss = SpreadsheetApp.openById('1QEAFU_Mlcl-UkOdynvD_x_1lxUgk6Hr8ovuLjIZWuQ0');
  var respSheet = ss.getSheetByName('Form Responses 1');
  
  //create spreadsheet variables - a range containing all RESPONSE IDs and EVENT IDs
  var textFinder = respSheet.createTextFinder('Response ID');
  var respIdCol = textFinder.findNext().getColumn();
  
  var textFinder2 = respSheet.createTextFinder('Event ID');
  var evIdCol = textFinder2.findNext().getColumn();
    
  var bothIds = respSheet.getRange(1, respIdCol, respSheet.getLastRow(), 2).getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
  
  //create spreadsheet variables - 1) array of RESPONSE IDs for all submissions and 2) array of EVENT IDs for all submissions
  var arrRespId = [];
  var arrEvId = [];
  
  bothIds.forEach(function(row) {
    var rVal = row[0];
    var eVal = row[1];
    arrRespId.push(rVal);
    arrEvId.push(eVal);
  });
  Logger.log('Submission Response ID: ' + respID);
  
  
  /**************************************************************************************************
   *      [1] Read and Store Item Response Values; Create Various Variables for Use in Script       *
   **************************************************************************************************/

  //[1.1] Create item response variables and store response values
    
    for (var i=0; i < itemResponses.length; i++) {
      switch (itemResponses[i].getItem().getTitle()) {   
        case 'Staff':
          var staff = itemResponses[i].getResponse();
          break;
        case 'Event Date':
          var evDate = itemResponses[i].getResponse();
          break;
        case 'Event Start Time':
          var startTime = itemResponses[i].getResponse();
          break;
        case 'Event Duration (HH:MM)':
          var evDuration = itemResponses[i].getResponse();
          break;
        case 'Event Title':
          var evTitle = itemResponses[i].getResponse();
          break;
        case 'Address':
          var evAddress = itemResponses[i].getResponse();
          break;
        case 'ZIP Code':
          var zipAndBoro = itemResponses[i].getResponse();
          var zipLength = zipAndBoro.length;
          var zip = zipAndBoro.substring(0,5);
          var boro = zipAndBoro.substring(8, zipLength);
          break;
        case 'Event Description':
          var evDescrip = itemResponses[i].getResponse();
          break;
        case 'Event Status':
          var evStatus = itemResponses[i].getResponse();
          break;
        case 'Expected Number of Attendees':
          var numAttendees = itemResponses[i].getResponse();
          break;
        case 'Additional Staffing Needs (Canvassers or Additional Campaign Team)':
          var staffNeedsNotes = itemResponses[i].getResponse();
          break;
        case 'Please list any additional resources needed to execute this event (number of mask, palm cards, languages)':
          var otherNeedsNotes = itemResponses[i].getResponse();
          break;
        case 'Do you need support from the following resources?':
          var resourceNeeds = itemResponses[i].getResponse();
          break;
        case 'Community Partners':
          var evPartners = itemResponses[i].getResponse();
          break;
        case 'Event Type':
          var evType = itemResponses[i].getResponse();
          break;
        case 'Initiative':
          var init = itemResponses[i].getResponse();
          break;
        case 'Target Market':
          var mkt = itemResponses[i].getResponse();
          break;
      } 
    }
  
  
  //[1.2] Create and select calendars where event will be created
  
    //create borough calendar ID variables and citywide calendar variable
    var xcal = 'c_vv621qlsub3diuthdklpil7r6g@group.calendar.google.com';
    var kcal = 'c_lpn9j47k7ua5n5m50a51ktjiv4@group.calendar.google.com';
    var mcal = 'c_8n9epegtgabgd84igltb4h5r2o@group.calendar.google.com';
    var qcal = 'c_t60ecpemls30u06harn7fcr9vs@group.calendar.google.com';
    var rcal = 'c_802kodj75g6vqq62pkt27qveb0@group.calendar.google.com';
    var NYCcal = 'c_9uqjt9n5tmmbq1jqbocc4s9go4@group.calendar.google.com';
  
    //select the appropriate borough calendar
    switch (boro) {
      case 'Bronx':
        var cal = xcal;
        break;
      case 'Brooklyn':
        var cal = kcal;
        break;
      case 'Manhattan':
        var cal = mcal;
        break;
      case 'Queens':
        var cal = qcal;
        break;
      case 'Staten Island':
        var cal = rcal;
        break;    
    }
  
  
  //[1.3] Clear any existing partner entries in the partner sheet (to be later (re)populated with any partners selected)
    var partnerSheet = ss.getSheetByName('Events by Partner');
  
    //create column index for response ID
    var respFinder = partnerSheet.createTextFinder('Form Response ID');
    var partnerRespIdCol = respFinder.findNext().getColumn();
    
    //check if there are any existing events in the sheet associated with this response ID
    var respIDs = partnerSheet.getRange(1, partnerRespIdCol, partnerSheet.getLastRow()).getValues();
    
    for (var row = 0; row < partnerSheet.getLastRow(); row++) {
      var val = respIDs[row];
      if (val==respID) {
        var partnersExist = true;
        break;
      } 
      else if (val!==respID) {
        var partnersExist = false;
      }
    }

    Logger.log('Partners Exist: ' + partnersExist);

    //if there are any existing events in the sheet associated with this response ID, delete them
    if (partnersExist == true) {
      for (var row = 1; row <= partnerSheet.getLastRow();) {
        var val = partnerSheet.getRange(row, partnerRespIdCol).getValue();
        if (val==respID) {
          partnerSheet.deleteRow(row);
        } else {
          row++;
        }
      }    
    }
    
  
  //[1.4] Check if submission is to cancel an event; if yes, delete event and end execution
    
    //check if submission is editing an existing event or creating a new event
    for (var row=0; row < respSheet.getLastRow(); row++) {
      var val = arrRespId[row];
      if (val==respID) {
        var newEntry = false;
        break;
      } 
      else if (val!==respID) {
        var newEntry = true;
      }
    }
         
    //for existing events, get the event ID
    if (newEntry == false) {
      for (var row = 0; row < respSheet.getLastRow(); row++) {
        var checkResp = bothIds[row];
        if (checkResp[0] == respID) {
          var oldEvId = checkResp[1];
          break;
        }
      }    
    }
  
    Logger.log('New Entry: ' + newEntry);
    Logger.log('First Event Status: ' + evStatus);
  
    //if Event Status is 'Canceled'...
    if (evStatus == 'Canceled') {
    
      //...and the submission is a new entry
      if (newEntry == true) {
        
        //then log response ID
        var respIdCell = respSheet.getRange(respSheet.getLastRow(), respIdCol);
        respIdCell.setValue(formResponse.getId());
        
        //and log custom message in event ID field
        var evIdCell = respSheet.getRange(respSheet.getLastRow(), evIdCol);
        evIdCell.setValue('This event was submitted as canceled. No event was created.');
        Logger.log('The event was successfully not created.');
        
        return;
      
      //...and the submission is not a new entry
      } else if (newEntry == false) {
        
        //then cancel the event
        Calendar.Events.remove(cal, oldEvId, {sendUpdates: 'all'});
        Calendar.Events.remove(NYCcal, oldEvId);
        return;
      }
    }
        

  //[1.5] Read in all user names and email addresses, then store only attendee data 
  
    //pull user name and email address from user list spreadsheet, trim whitespace
    var userSheet = ss.getSheetByName('User List');
  
    var userData = userSheet.getRange('A1').getDataRegion().getValues();
    userData.forEach(function(arr, i) {
      arr.forEach(function(val, j) {
        arr[j] = val.replace(/\s\s+/g, ' ').trim();
      });  
      userData[i] = arr
    });
  
    //filter resulting two-dimensional array by name using list of staff attending
    var attData = userData.filter(function(row) {
      return staff.includes(row[0]);
    });
      
 
  /**************************************************************************************************
   *                               [2] Create and Send Event                                        *
   **************************************************************************************************/
    
  //[2.1] Create and set event parameters
  
    //create year variable, infer event year, then store year value
    // NOTE: if event month precedes form submission month, event year is assumed to be the following calendar year
    //  e.g. 1 [i.e. Jan] event < 12 [i.e. Dec] submission ---> event will be held NEXT YEAR
    //  e.g. 3 [i.e. Mar] event !< 3 [i.e. Mar] submission ---> event will be held THIS YEAR
    //  e.g. 4 [i.e. Apr] event !< 3 [i.e. Mar] submission ---> event will be held THIS YEAR
     var evYear;
     var yrCheck = new Date(evDate);
     var today  = new Date();
     if (yrCheck.getMonth() < today.getMonth()) {
       evYear = (today.getFullYear() + 1).toString();
     } else {
       evYear = today.getFullYear().toString();
     }
  
    //create variables to use to create an endTime variable below
    var hours = evDuration.split(':');
    var seconds = (+hours[0])*60*60 +  (+hours[1])*60;
    
    //create date and location variables and store date and location values
    var startDT = new Date(evYear + '-' + evDate + 'T' + startTime);
    var endDT = new Date(startDT.getTime());
    endDT.setSeconds(endDT.getSeconds() + seconds);
    var evLocation = evAddress + '\n' + boro + ', NY ' + zip;
  
    //create an array to store email address objects for 1) respondent and 2) all attendees
    var evGuests = [{email: formResponse.getRespondentEmail().toString()}]
    attData.forEach(function(row) {
      var attEmail = row[2];
      evGuests.push({email: attEmail.toString()})
    });
    evGuests = [...new Set(evGuests)];
    Logger.log(evGuests);
  
  
    
  //[2.2] Select appropriate event color, create event color variable, store value
  
    //pull ZIP code data (i.e. color key) for event invite color coding
    var zipSheet = ss.getSheetByName('ZIPs');
  
    var textFinder3 = zipSheet.getDataRange().createTextFinder('ZIP (Sorted)');
    var zipStart = textFinder3.findNext().getA1Notation();
    var zipData = zipSheet.getRange(zipStart).getDataRegion().getValues();
  
    //create a zipColor variable and store the event zip code's color name 
    for (var i=0; i<zipData.length; i++) {
      var j = zipData[i];
      if (j.includes(zipAndBoro)==true) {
        var zipColor = j[3];
        break;
      }
    }
  
    //recode zipColor to be readable for event creation 
    switch (zipColor) {
      case 'Blue':
        zipColor = 7;
        break;
      case 'Green':
        zipColor = 2;
        break;
      case 'Yellow':
        zipColor = 5;
        break;
      case 'Orange':
        zipColor = 6;
        break;
      case 'Red':
        zipColor = 11;
        break;
      default:
        zipColor = 1;
    }
  
        
  //[2.3] Create calendar event in borough and citywide calendars and send to attendees
    
    //create event resource variable using set parameters
    var event = {
      summary: evTitle,
      location: evLocation,
      description: evDescrip + '\n\nEvent Status: ' + evStatus + '\n\nUpdate Event: \n' + respUrl,
      start: {dateTime: startDT.toISOString()},
      end: {dateTime: endDT.toISOString()},
      attendees: evGuests,
      colorId: zipColor
    };

    Logger.log('Event ID 1: ' + oldEvId);

    //update event if event ID exists and create it if not
    if (newEntry == false) {
      Logger.log('Calendar cal is: ' + cal);
      
      try {
        event = Calendar.Events.patch(event, cal, oldEvId, {sendUpdates: 'all'});
      }
      catch(e1) {
        var calVals = [xcal, kcal, mcal, qcal, rcal];
        calVals.forEach(function(calVal){
          try {
            var oldEvent = Calendar.Events.move(calVal, oldEvId, cal);
          }
          catch(e2) {
            Logger.log('There was an error when trying to move from cal: ' + calVal);
          }  
        });
      }
      
      event = Calendar.Events.patch(event, cal, oldEvId, {sendUpdates: 'all'});
      event = Calendar.Events.patch(event, NYCcal, oldEvId);   
      
    } else if (newEntry == true) {
      event = Calendar.Events.insert(event, cal, {sendUpdates: 'all'});
      event = Calendar.Events.insert(event, NYCcal);      
    }

    //add event to borough and citywide calendars and send event invitations
    
    Logger.log('Event ID 2: ' + event.id);
    Logger.log('Form Response ID: ' + formResponse.getId());


  /**************************************************************************************************
   *                    [3] Log Selected Values to Responses Spreadsheet                            *
   **************************************************************************************************/

  //[3.1] FOR NEW SUBMISSION - Save form response ID and event ID to form responses spreadsheet
      
    //log response ID
    var respIdCell = respSheet.getRange(respSheet.getLastRow(), respIdCol);
    respIdCell.setValue(formResponse.getId());
      
    //log event ID
    var evIdCell = respSheet.getRange(respSheet.getLastRow(), evIdCol);
    evIdCell.setValue(event.id);

    //log edit URL
    var textFinder4 = respSheet.createTextFinder('Edit Response Url');
    var urlCol = textFinder4.findNext().getColumn();
    var urlCell = respSheet.getRange(respSheet.getLastRow(), urlCol);
    urlCell.setValue(respUrl);
    
    
  /**************************************************************************************************
   *                     [4] Log Community Partners Data to a Separate Sheet                        *
   **************************************************************************************************/

    //create column indices for each field by name (in case columns are ever moved around)
    
    var partnerFields = [
      {'field': 'Community Partner', 'val': null, 'col': null},
      {'field': 'Event Date', 'val': evDate, 'col': null},
      {'field': 'Event Start Time', 'val': startTime, 'col': null},
      {'field': 'Event Title', 'val': evTitle, 'col': null},
      {'field': 'Address', 'val': evAddress, 'col': null},
      {'field': 'Borough', 'val': boro, 'col': null},
      {'field': 'ZIP Code','val': zip, 'col': null}, 
      {'field': 'Event Description', 'val': evDescrip, 'col': null},
      {'field': 'Event Type', 'val': evType, 'col': null},
      {'field': 'Initiative', 'val': init, 'col': null},
      {'field': 'Target Market', 'val': mkt, 'col': null}
    ];

    partnerFields.forEach(function(obj) {
      var finder = partnerSheet.createTextFinder(obj.field);
      var colVar = finder.findNext().getColumn();            
      obj.col = colVar;
    });
      
    //if any partners were selected
    if (evPartners != undefined) {
      //then for each partner selected in the submission... 
      evPartners.forEach(function(partner) {  
        var row = partnerSheet.getLastRow();
        row++;
        
        //save the partner name as the value for the 'val' property of the 'Community Partner' field object
        var partnerObj = partnerFields[0];
        partnerObj.val = partner;
        
        //log the values of all event characteristics to their respective columns
        partnerFields.forEach(function(field) {
          var cell = partnerSheet.getRange(row, field.col);
          cell.setValue(field.val);
        });
        
        //log the form response ID in the appropriate column
        var respIdCell = partnerSheet.getRange(row, partnerRespIdCol);
        respIdCell.setValue(formResponse.getId());
      });
    }

                             // DISABLED UNTIL POTENTIAL APPROVAL
/****************************************************************************************************

   **************************************************************************************************
   *                  [5] Send Email Alerts According to Event's Resource Needs                     *
   **************************************************************************************************

  //[5.1] Pull and store email addresses for recipients to be emailed if resources are needed
    
    //pull email addresses and create an array to store
    var textFinder = userSheet.getDataRange().createTextFinder('Any Support');
    var alertEmailStart = textFinder.findNext().getA1Notation();
    var alertEmails = userSheet.getRange(alertEmailStart).getDataRegion().getValues();
    alertEmails.forEach(function(arr, i) {
      arr.forEach(function(val, j) {
        arr[j] = val.replace(/\s\s+/g, ' ').trim();
      });  
      alertEmails[i] = arr
    });
  
    var anyNeedsEmails = alertEmails.map(function(value,index) { return value[0]; });
    anyNeedsEmails.shift();
    anyNeedsEmails = anyNeedsEmails.filter(Boolean);
  
    var CITneedsEmails = alertEmails.map(function(value,index) { return value[1]; });
    CITneedsEmails.shift();
    CITneedsEmails = CITneedsEmails.filter(Boolean);
 
    var specificNeeds = ['Communications','Intergovernmental Affairs', 'Testing Site']
  
    //select email recipients
    if (resourceNeeds != null) {
      if (resourceNeeds.some(i => specificNeeds.includes(i))) {
      var recipients = CITneedsEmails.toString();
    } else {
      var recipients = anyNeedsEmails.toString();        
    }
     
      
  //[5.2] Set email parameters and send to recipients
      
    //set email parameters
    resourceNeeds = resourceNeeds.join('<br>-');

    var replyAddress = formResponse.getRespondentEmail();
    var subject = evTitle + ' - Resources Needed';
    var body = 'This event will require support from the following resources: <br><br>' 
                + '-' + resourceNeeds
                + '<br><br><b><u>Event Details</u>: </b>'
                + '<br><b>Title: </b>' + evTitle
                + '<br><b>Date and Time: </b>' 
                + startDT.toLocaleDateString('en-US', {weekday: 'short', month: 'numeric', day: 'numeric', year: '2-digit'}) + ', ' 
                + startDT.toLocaleTimeString('en-US', {hour: 'numeric', minute: 'numeric'}) + ' - ' 
                + endDT.toLocaleTimeString('en-US', {hour: 'numeric', minute: 'numeric'})
                + '<br><b>Location: </b>' + evAddress + ', ' + boro + ', NY ' + zip
                + '<br><b>Expected Number of Attendees: </b>' + numAttendees
                + '<br><b>Description: </b>' + evDescrip
                + '<br><br><b>Additional Staffing Needs: </b><br>'
                + staffNeedsNotes
                + '<br><br><b>Additional Resources Needed: </b><br>'
                + otherNeedsNotes + '<br>';  
         
    //send email
    MailApp.sendEmail({
      to: recipients,
      replyTo: replyAddress,
      subject: subject,
      htmlBody: body
    }); 
    
  }

********************************************************************************************************/


  /**************************************************************************************************
   *                                     END SCRIPT                                                 *
   **************************************************************************************************/
         
}