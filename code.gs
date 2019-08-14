/** 
 * Google app script to interact with Response sheet for EMATS system.
 *
 * This script has functions that do the following operations: 
 * - Generate a Unique ID for each participant upon successful registration
 * - Wrapper for transforming the UID to a QR Code
 * - Auto-email response to send the corresponding QR code to participants
 * - Attendance Tracking function to validate QR code when presented by participant at the event
 * - Pie-chart for demographics and other useful features
 *
 * Written by : EMATS Dev Team, 2019
 * 
 **/

function doGet(e) {
  
  //MODIFY: Put the link of the spreadsheet
  var ss = SpreadsheetApp.openByUrl("Spreadsheet Link");
  var sheet = ss.getSheetByName("Responses");
  
  var colindex = getByName("Attendance", sheet);  
  var values = sheet.getDataRange().getValues();
  
  
  var attendingcount = 0;
  var attendingtotal = values.length-1;

  for (var i =1; i < values.length; i++){
    if(values[i][colindex] == "Yes"){    
      attendingcount++;
    }
    }
  
  var data = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, 'Turnout')
      .addColumn(Charts.ColumnType.NUMBER, 'Percentage')
      
      .addRow(['Attended', attendingcount])
      .addRow(['Not Attended', attendingtotal-attendingcount])
      .build();

  var chart = Charts.newPieChart()
      .setDataTable(data) 
      .setTitle('Attendance')
      .build();


  var htmlOutput = HtmlService.createHtmlOutput().setTitle('Event Charts');
  
  var imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes());
  var imageUrl = "data:image/png;base64," + encodeURI(imageData);
  htmlOutput.append("Attendance Chart: <br/>");
  htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");
  return htmlOutput;
 
  
}


// This is what the webservice does when it receives a post request. It does so when the app scans a QR code.
function doPost(e){
  
  //MODIFY: Put the link of the spreadsheet
  var ss = SpreadsheetApp.openByUrl("Spreadsheet Link");
  var sheet = ss.getSheetByName("Responses");
  var sdata = e.parameter.sdata;
  
  //Search the generated codes, find which row the scanned data is.
  var row = rowOf(sdata,getByName("QR Code",sheet), sheet); 
  
  //The cell to be updated, in the same row and "Attendance" column.   
  if (row !== -1){
  
    var cell = sheet.getRange(row,getByName("Attendance",sheet)+1,1,1);
    cell.setValue("Yes");
    
  }   

}


// Send an e-mail every hour to those we did not send yet. Then set the flag to sent.

function SendMail(){
  
  //MODIFY: Put the link of the spreadsheet
  var ss = SpreadsheetApp.openByUrl("Spreadsheet Link");
  var sheet = ss.getSheetByName("Responses"); 
  
  
  var values = sheet.getDataRange().getValues();
  var colindex = getByName("Email Address", sheet);
  
  var adresses = [];
  
  for (var i =1; i < values.length; i++){
    adresses.push(values[i][colindex]);
    }  
  
  
  for each (var address in adresses){
  
    var row = rowOf(address,getByName("Email Address",sheet), sheet);
    var flagcell = sheet.getRange(row,getByName("Email Status",sheet)+1,1,1);
    var flag = flagcell.getValue();       
    
    if (flag !== "Sent"){
      
      
      var name = sheet.getRange(row,getByName("Name",sheet)+1,1,1).getValue();
      var QRcell = sheet.getRange(row,getByName("QR Code",sheet)+1,1,1);
      var eventCode = sheet.getRange(row,getByName("Event Code",sheet)+1,1,1).getValue();
      var codeData = new_id();
      
      QRcell.setValue(codeData);

      var codeUrl = "https://chart.googleapis.com/chart?chs=250x250&cht=qr&chl=" + codeData;
      var codeBlob = UrlFetchApp
                         .fetch(codeUrl)
                         .getBlob()
                         .setName("codeBlob");
    
      var logoUrl = "https://drive.google.com/uc?export=view&id=16uKbBsXO-1Babif0NhE9YOy9_RSwTda7";
      var logoBlob = UrlFetchApp
                         .fetch(logoUrl)
                         .getBlob()
                         .setName("logoBlob");     
      
      var subject = 'Registration confirmation for Event ID ' + eventCode.toString();
      
      
      MailApp.sendEmail({
        to: address.toString(),
        subject: subject,
        htmlBody: "Thank you for registering, "+ name +"! <br>" +
        "<br>" +
        "Here is your QR code, please show it at the event venue: <br>" +
        "<br>" +
        "<img src='cid:code'> <br>" + 
        "Thank you, <br>" +
        "<br>" +
        "<img src='cid:logo' style='width:100px; height:88px;'/> <br>" +
        "EMATS Dev Team",
        inlineImages:
        {
        code: codeBlob,
        logo: logoBlob
      }
  });    
      
      flagcell.setValue("Sent");
    }      
  
  }
}



//Helper functions


//Get the row number of a data value, searched in a column number. The value should be in "", while the column index is an integer.
//Returns the row number of containingValue in columnIndex. If not found in there, returns -1.
function rowOf(containingValue, columnToLookInIndex, sheet) {
 
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var outRow;

  for (var i = 0; i < values.length; i++)
  {
    if (values[i][columnToLookInIndex] == containingValue)
    {
      outRow = i+1;
      break;
    }
    
    if (i == values.length - 1){
    
      outRow = -1;
    }
    
  }

  return outRow;
}

//Get the index of column, when the header is entered.
function getByName(colName,sheet) {
  //var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  return col;
}

// Function to generate random numbers in range (min, max)

function genRndNum(min,max){
  return Math.floor(Math.random()*(max-min + 1) + min);
}

function new_id(){
  var rndNum = genRndNum(1,999999999999);
  var numDigits = 10
  var id = ('0000' + rndNum).slice(-numDigits); //get the last N digits
  id = 'UID' + id; 
  return id;
}

