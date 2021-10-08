function scrapeDataDict() {
  
  var doc = DocumentApp.openById('1lU3CI-zVsTu9QmDab_TfmBlmYgfycsly6bNvNFo7HwA');
  var body = doc.getBody();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dictSheet = ss.getSheetByName('Testing_v1');
  var recodeSheet = ss.getSheetByName('_recode');
  var propsSheet = ss.getSheetByName('_properties');
  propsSheet.clearContents().appendRow(['Properties']);
   
  //create an array to store all field objects   
  //var fieldArray = [];
  
  //if the dictionary sheet is blank then add base fields as column headers
  if (dictSheet.getRange(1,1,1).isBlank()) {
    var propertyList = ['Heading', 'API Name', 'Field Name', 'Field Type', 'Object', 'Notes'];
    dictSheet.appendRow(propertyList);
  }
  
  //create and begin an array to store all unique properties and store in properties sheet
  var propertyList = dictSheet.getRange(1,1,1, dictSheet.getLastColumn()).getValues()[0];
  propertyList.map(prop => propsSheet.appendRow([prop]));
    
  scrapeFromBody(body);
    
/**********************************************************************************************************************************************************************/

  //collect all fields in the body of the doc and metadata for each field
  function scrapeFromBody(body) {

    //collect all horizontal rules in the doc
    var hr = null;
    var hrArray = [];
    var indexObjs = [];
   
    while (true) {
      hr = body.findElement(DocumentApp.ElementType.HORIZONTAL_RULE, hr);
      if (hr == null) break;
      hrArray.push(hr);
      var parent = hr.getElement().getParent();
      var parentIndex = body.getChildIndex(parent);
      indexObjs.push({'hrIndex': hr, 'parentIndex': parentIndex});
    }  
    
    for (a=0; a<5; a++) {
      var obj = indexObjs[a];
      Logger.log('parentIndex: ' + obj.parentIndex);
    }
    
    
    Logger.log('# of horizontal rules: ' + hrArray.length);
    
    var startField = 854;
            
    //get the document's first heading
    var heading = getFirstHeading(hrArray, indexObjs, startField, body);
    Logger.log('heading: ' + heading);
        
    
    //for each horizontal rule...
    for (i=startField; i < hrArray.length; i++) {
      
      //initialize nextHeading variable for later use
      var nextHeading = heading;
      
      //create a range containing all elements between that horizontal rule and the next one 
      var rangeBuilder = doc.newRange();
      var lastElement = null;
      
      
      //account for final field which has no end horizontal rule 
      if (i == hrArray.length - 1) {
        var element = hrArray[i].getElement();
        while (true) {
          lastElement = element;
          element = element.getNextSibling();
          if (element == null) {
            break;
          }                    
        }
        rangeBuilder.addElementsBetween(hrArray[i].getElement(), lastElement);
      } else {
      //all other fields
        rangeBuilder.addElementsBetween(hrArray[i].getElement(), hrArray[i+1].getElement());
      }
      
      var range = rangeBuilder.build();
      
      //get all elements in the rnage
      var rawElements = range.getRangeElements();
      
      //if the range contains a table, then call scrapeFromTables(table) and move on to next field
      var tableEls = rawElements.filter(el => (el.getElement().getType() == DocumentApp.ElementType.TABLE));
      if (tableEls.length > 0) {
        for (j=0; j < rawElements.length; j++) {
          if (isHeading(rawElements[j])) {
            nextHeading = updateHeading(rawElements[j]);
          } else if (rawElements[j].getElement().getType() == DocumentApp.ElementType.TABLE) {
            scrapeFromTables(rawElements[j].getElement(), heading);
           }            
        }
        heading = nextHeading;
        continue;        
      }

      
      //keep only PARAGRAPH-type and TEXT-type elements
      var textType = DocumentApp.ElementType.TEXT;
      var paraType = DocumentApp.ElementType.PARAGRAPH;
      var elements = rawElements.filter(el => (el.getElement().getType() == textType) || el.getElement().getType() == paraType);
      
      //check that the range contains non-empty PARAGRAPH-type elements
      var allParaEls = elements.filter(el => (el.getElement().getType() == paraType));
      var nonEmptyParaEls = allParaEls.filter(el => (el.getElement().asText().getText() != ''));
      
      //if the range does not contain non-empty PARAGRAPH-type elements, then check for (and update) heading, then skip and move to next field
      if (nonEmptyParaEls.length == 0) {
        for (j=0; j < elements.length; j++) {
          if (isHeading(elements[j])) {
            nextHeading = updateHeading(elements[j]);
          }
        }
        continue;        
      }
                
      //Logger.log('index ' + i + ' elements.length: ' + elements.length);
      //Logger.log('index ' + i + ' elements TEXT and HEADINGS:');
      //elements.forEach(function (el) {
      //  Logger.log('Element type: ' + el.getElement().getType());
      //  Logger.log('elText: ' + el.getElement().asText().getText() + '\nelHeading: ' + el.getElement().asParagraph().getHeading());
      //});
      
      //create an object to later store the field's metadata
      var fObj = {};
           
      //for each element in the range...
      for (j=0; j < elements.length; j++) {
        
        //check if the element is a heading, if yes, update heading variable, then skip the rest of the loop
        if (isHeading(elements[j])) {
          nextHeading = updateHeading(elements[j]);
          continue;
        }
        
        //get the text of the element
        var elText = elements[j].getElement().asText().getText().trim();
                
        //get the text of the element if it contains ':'
        if (!elText.includes(':') || elText.charAt(0) == '*') {
          continue;        
        }
        
        //parse the text into a property name and value
        var [elProp, elValue] = parseText(elText);
        
        //if the property is not already in the list of properties, then add it
        if (!propertyList.includes(elProp)) { 
          addProperty(elProp); 
        }
        
        //save the property and its value to the field object
        fObj[elProp] = elValue;
      }            
      
      //write the field object values to the dictionary
      var rowArray = [];
      
      //add the heading to the object
      fObj['Heading'] = heading;      
      
      if (!Object.keys(fObj).includes('API Name')) {
        //if field does not have 'API Name' as a property, add it to the object with value '.'
        fObj['API Name'] = '.';        
      }
      
      propertyList.forEach(function(prop) {
        var cellVal = null
        if (Object.keys(fObj).includes(prop)) {
          cellVal = fObj[prop];
        }
        rowArray.push(cellVal);
      });
      
      dictSheet.appendRow(rowArray);
      
      heading = nextHeading;
      
      //save the field object to the field array
      //fieldArray.push(fObj);
    }

    //print the number of properties, list of properties, and number of fields
    Logger.log('Number of properties: ' + propertyList.length);
    Logger.log('List of properties: ' + propertyList);
    //Logger.log('Number of fields: ' + fieldArray.length);
       
  }

  
  
/**********************************************************************************************************************************************************************/
   
  function scrapeFromTables(table, heading) {
    var numRows = table.getNumRows();
    
    for (row = 1; row < numRows; row++) {
      var apiName = table.getCell(row, 2).getText();
      var fName = table.getCell(row, 1).getText();
      var notes = '';
      var fType = table.getCell(row, 3).getText();
      var object = table.getCell(row, 0).getText();
      
      var fObj = {'Heading': heading, 'API Name': apiName, 'Field Name': fName, 'Notes': notes, 'Field Type': fType, 'Object': object};
      var rowArray = [];
      
      propertyList.forEach(function(prop) {
        var cellVal = null
        if (Object.keys(fObj).includes(prop)) {
          cellVal = fObj[prop];
        }
        rowArray.push(cellVal);
      });
      
      dictSheet.appendRow(rowArray);     
    }
    return; 
  }  
  
/**********************************************************************************************************************************************************************/   

  function parseText(text) {
    //split text into property name and value
    text = text.split(':');
    var prop = text[0].trim();
    text.shift();
    var val = text.join().trim();
    
    //correct typos/recode property name values as needed
    var recodeData = recodeSheet.getDataRange().getValues();
    var recodeDict = recodeData.map(pair => ({value: pair[0], recode: pair[1]}));
    var badVals = recodeDict.map(obj => obj.value);
    
    if (badVals.includes(prop)) {
      var fieldObj = recodeDict.find(obj => obj.value == prop);
      prop = fieldObj.recode;   
    }
    
    return [prop, val];
  }

/**********************************************************************************************************************************************************************/

  function addProperty(prop) {
    //add to the list
    propertyList.push(prop);
    
    //write to the properties sheet
    propsSheet.appendRow([prop]);
    
    //create a new column header in the dictionary
    var newColCell = dictSheet.getRange(1, dictSheet.getLastColumn() + 1);
    newColCell.setValue(prop);
  }
   
  /**********************************************************************************************************************************************************************/
  
  function getFirstHeading(array, objArray, start, body) {
    var headingText = '';
    
    var hr = array[start];
    var hrObj = objArray.find(obj => obj.hrIndex == hr);
    var index = hrObj.parentIndex;
    
    var el = body.getChild(index);
    var prev = el.getPreviousSibling();
    
    while (true) {
      
      if (prev.getType() == DocumentApp.ElementType.TEXT) {
        if (prev.asText().getText() != '') {
          if (prev.getAttributes().FONT_SIZE > 12) {
            headingText = prev.asText().getText().trim();
            break;
          }
        }
      } else if (prev.getType() == DocumentApp.ElementType.PARAGRAPH) {
        if (prev.asParagraph().getText() != '') {
          if (prev.asParagraph().getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
            headingText = prev.asParagraph().getText().trim();
            break;
          }
        }
      }
      
      prev = prev.getPreviousSibling();
    }
   
    Logger.log('headingText: ' + headingText);
    return headingText;
  }
    
/**********************************************************************************************************************************************************************/  

  function isHeading (element) {
    var headingCheck = false;
    if (element.getElement().getType() == DocumentApp.ElementType.TEXT) {
      if (element.getElement().getAttributes().FONT_SIZE > 12) {
        headingCheck = true;
      }
    } else if (element.getElement().getType() == DocumentApp.ElementType.PARAGRAPH) {
       if (element.getElement().asParagraph().getHeading() == DocumentApp.ParagraphHeading.HEADING1) {
        headingCheck = true;
      }
    }
    return headingCheck;
  }
  
/**********************************************************************************************************************************************************************/  
  
  function updateHeading(element) {
    var headingText = ''
    if (element.getElement().getType() == DocumentApp.ElementType.TEXT) {
      headingText = element.getElement().asText().getText().trim();
    } else if (element.getElement().getType() == DocumentApp.ElementType.PARAGRAPH) {
      headingText = element.getElement().asParagraph().getText().trim();
    }
    return headingText;  
  }
  
  
             
}
