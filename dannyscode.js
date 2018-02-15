function main() {
  //get the ui to communicate
  var ui = SpreadsheetApp.getUi();
  //get the spreatsheet
  var ss = SpreadsheetApp.getActive();
  //create the Overall sheet
  var s = ss.getSheetByName("Overall");
  if(s) {
    s.setName("garbage sheet name that will be deleted");
  }
  s = ss.insertSheet("Overall");
  //delete all the sheets but Overall
  var sheets = ss.getSheets();
  for(var i = 0; i < sheets.length; i ++) {
    if(sheets[i].getName() != "Overall") {
      ss.deleteSheet(sheets[i]);
    }
  }
  //ask how many graders there are
  var graders = parseInt(ui.prompt("How many graders are there?").getResponseText());
  var names = [];
  //get the names of the graders + create a sheet for them
  for(var i = 0; i < graders; i ++) {
    names.push(ui.prompt("Grader #" + (i + 1) + "'s name?").getResponseText());
    ss.insertSheet(names[i]);
  }
  //get the amount of difference to flag
  var diff = ui.prompt("This sheet will flag applicants who's highest score and lowest score\n\
                       differ by more than a set amount. What would you like this amount to be?").getResponseText();
  //fill the cells in the overall sheet
  var sheet = ss.getSheetByName("Overall");
  //create the headings of the sheet
  var headings = ["Applicant"];
  for(var i = 0; i < graders; i ++) {
    headings.push(names[i] + "'s Score");
  }
  headings.push("Average Score");
  sheet.getRange(1,1,1,headings.length).setValues([headings]);
  //fill in the rows of the sheet
  for(var i = 2; i < 330; i++) { //there are only 328 rows filled in b/c this is the socruit record
    var string;
    for(var j = 0; j < graders; j ++) {
      string = "=IF(OR(";
      for(var k = 0; k < graders - 1; k ++) {
        string = string.concat(names[k] + "!B" + i + " =\"\",");
      }
      string = string.concat(names[graders - 1] + "!B" + i + "=\"\")");
      sheet.getRange(i, 2 + j, 1, 1).setValue(string + ", \"\", " + names[j] + "!B" + i + ")");
    }
    Logger.log(2 + graders);
    sheet.getRange(i, 2 + graders, 1 , 1).setValue(string + ",\"\", AVERAGE(B" + i + ":" + columnToLetter(2 + graders - 1) + "" + i + "))");
  }
  
}
  
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


/*
This function renames the sheets for new Committee heads and Coords.
It is not really required but i felt it was a nice touch
*/
function renameSheets() {
  //get the ui to communicate
  var ui = SpreadsheetApp.getUi();
  //if the user tells us to clear the sheets we will delete all of the applicant names, scores and comments from this document
  var overwrite = verifyOverwrite(ui);
  //get the spreadsheet
  var ss = SpreadsheetApp.getActive();
  var cells = ["B10", "B11", "C11"];
  //get the three sheets we need
 var sheets = ss.getSheets();
  var j = 0;
  for(var i = 0; i < sheets.length; i++){
    if(sheets[i].getName() != "Overall" && sheets[i].getName() != "Instructions") {
      if(overwrite) {
        for(var k = 2; k < 330; k ++) {
          sheets[i].getRange(k, 2, 1, 1).setValue("");
          sheets[i].getRange(k, 3, 1, 1).setValue("");
        }
      }
      sheets[i].setName(ss.getSheetByName("Instructions").getRange(cells[j]).getValue());
      j++;
    }
  }
  
  //for some reason I have to regen the cells for google sheets to update the values in them.
  //Idk why and its dumb but this is the workaround
  fillStuffCauseImLazy(overwrite);
}

/*
You should not have to touch the stuff below. It was used in the creation of the sheet
sheets would not drag down the column and adjust the numbers because the cell was in ""
So I wrote this chunk to fill in 328 rows cause doing that by hand would have sucked
*/


function fillStuffCauseImLazy(overwrite) {
  //get the spreadsheet that we want to use
  var ss = SpreadsheetApp.getActive();
  //get the tab for the Data
  var sheet = ss.getSheetByName("Overall");
  for(var i = 2; i < 330; i ++) {
    if(overwrite) {
      sheet.getRange(i, 1, 1, 1).setValue("");
    }
    sheet.getRange(i, 2, 1, 1).setValue("=IF(OR(INDIRECT(\"'\"&Instructions!$B$11&\"'!B" + i + 
                                        "\") = \"\", INDIRECT(\"'\"&Instructions!$C$11&\"'!B" + i + 
                                        "\") = \"\"), \"\", INDIRECT(\"'\"&Instructions!$B$10&\"'!B" + i + "\"))");
    sheet.getRange(i, 3, 1, 1).setValue("=IF(OR(INDIRECT(\"'\"&Instructions!$B$10&\"'!B" + i + 
                                        "\") = \"\", INDIRECT(\"'\"&Instructions!$C$11&\"'!B" + i + 
                                        "\") = \"\"), \"\", INDIRECT(\"'\"&Instructions!$B$11&\"'!B" + i + "\"))");
    sheet.getRange(i, 4, 1, 1).setValue("=IF(OR(INDIRECT(\"'\"&Instructions!$B$10&\"'!B" + i + 
                                        "\") = \"\", INDIRECT(\"'\"&Instructions!$B$11&\"'!B" + i + 
                                        "\") = \"\"), \"\", INDIRECT(\"'\"&Instructions!$C$11&\"'!B" + i + "\"))");
  }
}


//Below is code I found online and modified to create an alert notifying a user that they are about to overwrite old data
//and asking them to confirm their actions
function verifyOverwrite(ui) {
  var result = ui.alert(
     'Do you want to clear out all of the cells in this sheet?, Clicking Yes will peremently delete all data (so save a backup)',
      ui.ButtonSet.YES_NO);
 
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    return 1;
    
  } else {
    // User clicked "No" or X in the title bar.
    return 0;
  }
}


