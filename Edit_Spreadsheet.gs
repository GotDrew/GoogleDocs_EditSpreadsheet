/*
    Script works with a Spreadsheet to find a Cell on different Sheets and change the value and colour.
    Needs:
        Spreadsheet ID
        Sheet number (done with Days here)
        Something to define the Rows and Columns (Room and Period)
        New Text (Class name)
        Optional Colout, change text colour with: setFontColor(#000000) //black
    
    Andrew Hopcroft
    hopcroft.andrew@gmail.com
*/
function onFormSubmit (e)
{
  //Send test email
  //MailApp.sendEmail("USER@SOMEWHERE.NET", "Init - Change Cell", "Started");
  
  //Get the responses
  var Responses = e.response.getItemResponses();
  var NameOfPBS = Responses[0].getResponse();
  var ClassName = Responses[1].getResponse();
  var DayName = Responses[2].getResponse();
  var PeriodNum = Responses[3].getResponse();
  var RoomName = Responses[4].getResponse();
  var ColourName = Responses[5].getResponse();
  
  //Change responses to cell Rows and Columns.
  if (PeriodNum == "1") var cellC = "C";
  else if (PeriodNum == "2") var cellC = "D";
  else if (PeriodNum == "3") var cellC = "E";
  else if (PeriodNum == "4") var cellC = "F";
  else if (PeriodNum == "5") var cellC = "G";
  else if (PeriodNum == "6") var cellC = "H";
  else if (PeriodNum == "7") var cellC = "I";
  else if (PeriodNum == "8") var cellC = "J";
  
  if (RoomName == "W109") var cellR = 3;
  else if (RoomName == "W110")  var cellR = 5;
  else if (RoomName == "W112")  var cellR = 7;
  else if (RoomName == "W113")  var cellR = 9;
  else if (RoomName == "W114")  var cellR = 11;
  else if (RoomName == "W115")  var cellR = 13;
  else if (RoomName == "W116")  var cellR = 15;
  else if (RoomName == "W218")  var cellR = 17;
  else if (RoomName == "W222")  var cellR = 19;
  else if (RoomName == "W226")  var cellR = 21;
  else if (RoomName == "W227")  var cellR = 23;
  
  //Grab a colour from responses
  if (ColourName == "CLEAR")  var cellColour = "#ffffff";
  else if (ColourName == "YR 6") var cellColour = "#980000";
  else if (ColourName == "YR 7")  var cellColour = "#9fc5e8";
  else if (ColourName == "YR 8")  var cellColour = "#0000ff";
  else if (ColourName == "YR 9")  var cellColour = "#ff0000";
  else if (ColourName == "YR 10")  var cellColour = "#00ff00";
  else if (ColourName == "Biology")  var cellColour = "#38761d";
  else if (ColourName == "Chemistry")  var cellColour = "#ffff00";
  else if (ColourName == "Phyiscs")  var cellColour = "#ff00ff";
  
  
  //Find the Document to edit
  var spreadsheet = SpreadsheetApp.openById(NameOfPBS);
  SpreadsheetApp.setActiveSpreadsheet(spreadsheet);
  
  
  //Find Sheets with corrent Day
  var numSheets = spreadsheet.getSheets();
  
  for (var i in numSheets) 
  {
    var theSheet = numSheets[i];               
    var sheetName = numSheets[i].getName();      //Get the name of the Sheet
    var shortName = sheetName.substring(0, 3);   //Make it short 
    shortName.toLowerCase();                     //Make it lower case
    var shortDay = DayName.substring(0,3);       //Make the responses day name Short
    shortDay.toLowerCase();                      //And lower case
    
    if( shortName == shortDay )                  //If they match then we find the cell and change its values and colour
    {
      var theCell = "";                          //error when done like: var theCell = cellC + cellR;
      theCell = theCell + cellC + cellR;         //no error when done like this.
      var editCell = theSheet.getRange(theCell); //Grabbing the correct cell
      editCell.setValue(ClassName.toString());   //Change Value
      editCell.setBackground(cellColour);        //Change Colout
    }
  }
  //Send finished email.
  //MailApp.sendEmail("USER@SOMEWHERE.NET", "Complete - Change Cell", "Finished.\n\n" + DayName + "\nRoom: " + RoomName + " During P" + PeriodNum+ "\nNew Class: " + ClassName.toString());
}
