
function addPoints() 
{

  var comNum = 0;
  var valid = 0;
  var ui = SpreadsheetApp.getUi();
 
  while (valid == 0)
  {
    var enterCommittee = ui.prompt('Enter Committee Name: ');
    var response = enterCommittee.getResponseText();
    Logger.log(response);
    if (response == 'Outreach')
    {
      comNum = 1;
      valid = 1;
    }
    else if (response == 'Community')
    {
      comNum = 2;
      valid = 1;
    }
    else if (response == 'MES')
    {
      comNum = 3;
      valid = 1;
    }
    else if (response == 'Fundraising')
    {
      comNum = 4;
      valid = 1;
    }
    else if (response == 'Recruitment')
    {
      comNum = 5;
      valid = 1;
    }
    else if (response == 'PL')
    {
      comNum = 6;
      valid = 1;
    }
    else if (response == 'Team Tech')
    {
      comNum = 7;
      valid = 1;
    }
    else if (response == 'Info/Mark')
    {
      comNum = 8;
      valid = 1;
    }
    else if (response == 'GradSWE')
    {
      comNum = 9;
      valid = 1;
    }
    else if (response == 'President')
    {
      comNum = 10;
      valid = 1;
    }
    else if (response == 'EVP')
    {
      comNum = 11;
      valid = 1;
    }
    else if (response == 'IVP')
    {
      comNum = 12;
      valid = 1;
    }
    else if (response == 'Secretary')
    {
      comNum = 13;
      valid = 1;
    }
    else if (response == 'Treasurer')
    {
      comNum = 14;
      valid = 1;
    }
    else if (response == 'NomCom')
    {
      comNum = 15;
      valid = 1;
    }
    else
    {
     // var invalid = ui.prompt('Invalid Name');
     var invalid = Browser.msgBox('Invalid Name')
      valid = 0;
    }
 }
  
  var enterEvent = ui.prompt('Enter Event Name:');
  var eventName = enterEvent.getResponseText();
  var enterPoint = ui.prompt('Enter Number of Points:');
  var pointVal = enterPoint.getResponseText();
  var enterURL = ui.prompt('Enter URL of attendance (Google Sheets):');
  var url = enterURL.getResponseText();
  var ss = SpreadsheetApp.openByUrl(url).getSheets()[0]; //event attendance

  
  var added = 0;  //prevent multiple additions of name to committeeList
  var addedM = 0;
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1G0vLDHH2i6LRKRb6Vrjusg4iWoCUX23A9_j3A-39rvU/edit#gid=0').getSheets()[comNum]; //Committee Point System sheet and test Team Tech sheet
  var masterSheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1G0vLDHH2i6LRKRb6Vrjusg4iWoCUX23A9_j3A-39rvU/edit#gid=0').getSheets()[0]; //master list sheet on point system spreadsheet
  var data = ss.getDataRange().getValues();  //data on event attendance sheet
  var committeeData = sheet.getDataRange().getValues(); //data on committee list
  var masterData = masterSheet.getDataRange().getValues();  //data on Master List
  sheet.insertColumnAfter(sheet.getLastColumn()); //insert new column for new event points
  var eventNameCell = sheet.getRange(1, sheet.getLastColumn()+1);
  eventNameCell.setValue(eventName); //set name of column as event name
  
  for (var i = 1; i < data.length; i++)
  {
    for (var j = 1; j < committeeData.length; j++) 
    {
      if (data[i][1] == committeeData[j][1])
      {
        //if name on attendance sheet is already on committeeList, just add a point value for the event column
        added = 1;
        var lastColumn=sheet.getLastColumn();
        var addValCell= sheet.getRange(j+1, lastColumn);
        Logger.log(addValCell);
        addValCell.setValue(pointVal);
      }
      
    }
    for (var j =0; j < masterData.length; j++)
    {
      if (data[i][1] == masterData[j][1])
      {
        //if name on attendance sheet is already on masterList, move on
        addedM = 1;
      }
    }
    //Logger.log(added); //test checkpoint
      if (added == 0)
      {
       //if name on attendance sheet is not on committeeList, add name, add formula for total sum, and add point value for event column
        sheet.appendRow([data[i][0], data[i][1]]); //pointVal]);
        var lastRow = sheet.getLastRow();
        var lastColumn = sheet.getLastColumn();
        var totalCell = sheet.getRange(lastRow, 3);
        
        lastIndex = sheet.getLastRow();
        totalCell.setFormula("=SUM(D" + lastIndex + ":DA" + lastIndex + ")");
        var addValCell = sheet.getRange(lastRow, lastColumn);
        addValCell.setValue(pointVal);
      }
    if (addedM == 0)
    {
      //if name on attendance sheet is not on masterList, add name, add formula for total sum, and add point value for event column
        masterSheet.appendRow([data[i][0], data[i][1]]);
        var lastMRow = masterSheet.getLastRow();
        var lastMColumn=masterSheet.getLastColumn();
        var totalMCell = masterSheet.getRange(lastMRow, 3);
        
        totalMCell.setFormula("sum(iferror(index('Info/Mark'!C:C,match(B" + lastMRow +",'Info/Mark'!B:B,0)),0),iferror(index(TeamTech!C:C,match(B" + lastMRow +",TeamTech!B:B,0)),0),IFERROR(index(Social!C:C,match(B" + lastMRow +",Social!B:B,0)),0),IFERROR(index(Outreach!C:C,match(B" + lastMRow +",Outreach!B:B,0)),0),IFERROR(index(Fundraising!C:C,match(B" + lastMRow +",Fundraising!B:B,0)),0),IFERROR(index(Community!C:C,match(B" + lastMRow +",Community!B:B,0)),0),IFERROR(index(Recruitment!C:C,match(B" + lastMRow +",Recruitment!B:B,0)),0),IFERROR(index(GradSWE!C:C,match(B" + lastMRow +",GradSWE!B:B,0)),0),IFERROR(index(IVP!C:C,match(B" + lastMRow +",IVP!B:B,0)),0),IFERROR(index(EVP!C:C,match(B" + lastMRow +",EVP!B:B,0)),0),IFERROR(index(Secretary!C:C,match(B" + lastMRow +",Secretary!B:B,0)),0),IFERROR(index(President!C:C,match(B" + lastMRow +",President!B:B,0)),0))");
        
       
     }
    added = 0 ; //reset added variable for next name on attendance sheet
    addedM = 0;
  }

}



