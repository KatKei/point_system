
function addName()
{
  var pointVal= 10;        //add user input
  var eventName = "Event"; //add user input
  var added = 0;  //prevent multiple additions of name
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1t0ESd6up31lkZzkFZl5H0vJkAR1Kwyr8ovP3yUBSVUk/edit#gid=635060399').getSheets()[0]; //event attendance (google sheets from google form with name and netID (right now it's url dependent, maybe user input?)
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/18aoYXJ99mSbO1XGwt4TDizwfjSDAtoW8m8270S9PdHo/edit#gid=17430624050').getSheets()[4]; //test Point System spreadsheet and test Team Tech sheet (url of entire spreadsheet with masterList & committee lists --only needs to be changed if link changes)
  var data = ss.getDataRange().getValues();  //data on event attendance sheet
  var masterData = sheet.getDataRange().getValues(); //data on masterList
  sheet.insertColumnAfter(sheet.getLastColumn()); //insert new column for new event points
  var eventNameCell = sheet.getRange(1, sheet.getLastColumn()+1);
  eventNameCell.setValue(eventName); //set name of column as event name

  for (var i = 1; i < data.length; i++)
  {
    for (var j = 1; j < masterData.length; j++)
    {
      if (data[i][1] == masterData[j][0])
      {
        //if name on attendance sheet is already on masterList, just add a point value for the event column
        added = 1;
        var lastColumn=sheet.getLastColumn();
        var addValCell= sheet.getRange(j+1, lastColumn);
        Logger.log(addValCell);
        addValCell.setValue(pointVal);
      }

    }
    //Logger.log(added); //test checkpoint
      if (added == 0)
      {
       //if name on attendance sheet is not on masterList, add name, add formula for total sum, and add point value for event column
        sheet.appendRow([data[i][1], data[i][2], pointVal]);
        var lastRow = sheet.getLastRow();
        var lastColumn = sheet.getLastColumn();
        var totalCell = sheet.getRange(lastRow, 3);

        lastIndex = sheet.getLastRow();
        totalCell.setFormula("=SUM(D" + lastIndex + ":DA" + lastIndex + ")");
        var addValCell = sheet.getRange(lastRow, lastColumn);
        addValCell.setValue(pointVal);

      }
    added = 0 ; //reset added variable for next name on attendance sheet
  }

}
