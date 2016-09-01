
function compareList() 
{

  var comNum = 0;
  var addedM = 0; //prevent multiple additions of name to masterList
  var masterSheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1G0vLDHH2i6LRKRb6Vrjusg4iWoCUX23A9_j3A-39rvU/edit#gid=0').getSheets()[0]; //master list sheet on point system spreadsheet
  var masterData = masterSheet.getDataRange().getValues();  //data on Master List
  
  
  ml = masterData.length;
  for (comNum = 1; comNum <16; comNum++)
  {
  
    Logger.log(comNum);
  var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1G0vLDHH2i6LRKRb6Vrjusg4iWoCUX23A9_j3A-39rvU/edit#gid=0').getSheets()[comNum]; //Committee Point System sheet
  var committeeData = sheet.getDataRange().getValues(); //data on current committee list
  var cl = committeeData.length;
  
 
    for (var i = 1; i < cl; i++)
    {
    
      for (var j =0; j < ml; j++)
      {
        if (committeeData[i][1] == masterData[j][1])
        {
          //if name on commiteeList is already on masterList, move on
        addedM = 1;
        }
      }
    
    if (addedM == 0)
    {
      //if name on attendance sheet is not on masterList, add name, add formula for total sum, and add point value for event column
        masterSheet.appendRow([committeeData[i][0], committeeData[i][1]]);
        var lastMRow = masterSheet.getLastRow();
        var lastMColumn=masterSheet.getLastColumn();
        var totalMCell = masterSheet.getRange(lastMRow, 3);
        
        totalMCell.setFormula("sum(iferror(index('Info/Mark'!C:C,match(B" + lastMRow +",'Info/Mark'!B:B,0)),0),iferror(index(TeamTech!C:C,match(B" + lastMRow +",TeamTech!B:B,0)),0),IFERROR(index(Social!C:C,match(B" + lastMRow +",Social!B:B,0)),0),IFERROR(index(Outreach!C:C,match(B" + lastMRow +",Outreach!B:B,0)),0),IFERROR(index(Fundraising!C:C,match(B" + lastMRow +",Fundraising!B:B,0)),0),IFERROR(index(Community!C:C,match(B" + lastMRow +",Community!B:B,0)),0),IFERROR(index(Recruitment!C:C,match(B" + lastMRow +",Recruitment!B:B,0)),0),IFERROR(index(GradSWE!C:C,match(B" + lastMRow +",GradSWE!B:B,0)),0),IFERROR(index(IVP!C:C,match(B" + lastMRow +",IVP!B:B,0)),0),IFERROR(index(EVP!C:C,match(B" + lastMRow +",EVP!B:B,0)),0),IFERROR(index(Secretary!C:C,match(B" + lastMRow +",Secretary!B:B,0)),0),IFERROR(index(President!C:C,match(B" + lastMRow +",President!B:B,0)),0))");
        
       
     }
    //reset added variable for next name on attendance sheet
    addedM = 0;
  }
  }

}



