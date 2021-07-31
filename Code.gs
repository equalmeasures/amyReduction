var thisSheet = SpreadsheetApp.getActiveSpreadsheet();
var mainTab = thisSheet.getSheetByName("Form Responses 1");
var workingSheet = thisSheet.getSheetByName("WorkingSheet");

//this function is for testing - it gives you a "Send" menu button in the sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Send")
    .addItem("Emails",'createTimeTable')
  .addToUi();
}

//this is the main function if you want to debug
function createTimeTable(){
  var data = mainTab.getDataRange().getValues();
  
  for(var i = 1; i < data.length; i++){
    if(data[i][6]==="X"){

    }else{
      
      workingSheet.clear();

      var weightPerMg = data[i][5]/data[i][2];
      

      workingSheet.getRange(1,1).setValue("Day");
      workingSheet.getRange(1,2).setValue("Date");
      workingSheet.getRange(1,3).setValue("Mg");
      workingSheet.getRange(1,4).setValue("Weight");

      workingSheet.getRange(2,1).setValue(1);
      workingSheet.getRange(2,2).setValue(data[1][4]);
      workingSheet.getRange(2,3).setValue(data[1][2]);
      workingSheet.getRange(2,4).setValue(data[1][5]);

      for(var j = 3; j < 102; j++){
        workingSheet.getRange(j,1).setValue(j-1);

        var lastDate = new Date(workingSheet.getRange(workingSheet.getLastRow()-1,2).getValue());
        workingSheet.getRange(j,2).setValue(new Date(lastDate.getTime()+(1000 * 60 * 60 * 24)));

        var lastMg = workingSheet.getRange(workingSheet.getLastRow()-1,3).getValue();
        workingSheet.getRange(j,3).setValue(lastMg-((lastMg*data[i][3])/100).toFixed(2));

        workingSheet.getRange(j,4).setValue(((workingSheet.getRange(j,3).getValue())*weightPerMg).toFixed(4));
      }

      sendEmail(i+1);
    }
  }
}

function sendEmail(row) {
  
      const headers = workingSheet.getRange(1,1,1,4).getDisplayValues();

      const day = headers[0][0];
      const date = headers[0][1];
      const mg = headers[0][2];
      const weight = headers[0][3];

      const lastRow = workingSheet.getLastRow();
      const tablerangeValue = workingSheet.getRange(2,1,lastRow-1,4).getDisplayValues();

      const htmlTemplate = HtmlService.createTemplateFromFile('emailTable');

      htmlTemplate.day = day;
      htmlTemplate.date = date;
      htmlTemplate.mg = mg;
      htmlTemplate.weight = weight;

      htmlTemplate.tablerangeValue = tablerangeValue;

      const htmlForEmail = htmlTemplate.evaluate().getContent();

      MailApp.sendEmail({
        to:'equalmeasures99@gmail.com',
        subject: 'Your Time Table',
        htmlBody: htmlForEmail
      });

      
      mainTab.getRange(row,7).setValue("X");
      workingSheet.clear();
}

