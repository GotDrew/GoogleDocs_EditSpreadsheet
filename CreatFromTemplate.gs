/*
    Script that creates a new Speadsheet by copying a template file and filling it with data from other templates.
    Example: School Diary with subjects filled in.
    Attach script to Google Form with a submit trigger. Responses need to be ordered.
    Form Needs:
    	New File Name
        Start and End dates
    Also Needs:
    	Spreadsheet IDs for the templates
    
    Andrew Hopcroft
    hopcroft.andrew@gmail.com
*/

function onFormSubmit (e) 
{
	//Day name list
	var dayNames = new Array("Sunday", "Monday", "Tuesday","Wednesday", "Thursday", "Friday", "Saturday");
	var i = 0;
  
	//Make the new doc 
	var tmpFile = SpreadsheetApp.openById("FILL_IN_ID");
	var newSpread = tmpFile.copy("TempName");
	
	//Get the Template
	var tmpMonday = SpreadsheetApp.openById('FILL_IN_ID');
	var tmpTuesday = SpreadsheetApp.openById('FILL_IN_ID');
	var tmpWednesday = SpreadsheetApp.openById('FILL_IN_ID');
	var tmpThursday = SpreadsheetApp.openById('FILL_IN_ID');
	var tmpFriday = SpreadsheetApp.openById('FILL_IN_ID');
  
	try
	{
		var resNew = e.response.getItemResponses();
		
		//Get the First and Last Dates as Strings
		var rawFirstDate = resNew[1].getResponse();
		var rawLastDate = resNew[2].getResponse();
		
		//Fix the formatting of the Stings
		var splitFirstDate = rawFirstDate.split('-');
		var splitlastDate = rawLastDate.split('-');
		
		//Seting the Fixed Strings as Dates
		var whileToday = new Date();
		var whileLast = new Date();
		
		whileToday.setDate(splitFirstDate[2]);		//firstDate, 
		whileToday.setMonth(splitFirstDate[1]-1);	//Adjust Month
		whileToday.setFullYear(splitFirstDate[0]);	
		
		whileLast.setDate(splitlastDate[2]);		//lastDate, 
		whileLast.setMonth(splitlastDate[1]-1);		//Adjust Month
		whileLast.setFullYear(splitlastDate[0]);
		
		var adjustDay = whileToday.getDate();
		whileToday.setDate(adjustDay - 1);
		
		var nextDay = whileToday;
		
		//Get the number of days between the first and last
		var numberOfDays = Math.round(( Date.parse(whileLast) - Date.parse(whileToday) ) / 86400000);
		
		//While to add and edit the sheets
		while (i < numberOfDays) 
		{
			var dayOfMonth = whileToday.getDate();
			nextDay.setDate(dayOfMonth + 1);
			
			var dateText = dayNames[nextDay.getDay()] + " " + nextDay.getDate() + "/" + (nextDay.getMonth()+1) + "/" + nextDay.getFullYear();
			
			if (nextDay.getDay() != 0 && nextDay.getDay() != 6) 
			{
				//Get the correct day sheet
				if (nextDay.getDay() == 1) var sheet = tmpMonday.getSheets()[0];
				if (nextDay.getDay() == 2) var sheet = tmpTuesday.getSheets()[0];
				if (nextDay.getDay() == 3) var sheet = tmpWednesday.getSheets()[0];
				if (nextDay.getDay() == 4) var sheet = tmpThursday.getSheets()[0];
				if (nextDay.getDay() == 5) var sheet = tmpFriday.getSheets()[0];
				
				sheet.copyTo(newSpread).setName(dateText);        //Copy the Sheet
				
				newSpread.setActiveSheet(newSpread.getSheetByName(dateText));
				
				newSpread.getRange('b1').setValue(dateText);      //Change the Title
				
				var j = newSpread.getSheets();
				newSpread.moveActiveSheet(j.length);                   //Move it to the start
			}
			i = i+1;
		}
		
		newSpread.rename(resNew[0].getResponse());
	}
	catch(err)
	{
		//error
		MailApp.sendEmail("USER@SOMEWHERE.NET", "Error report", err.message);
	}
}
