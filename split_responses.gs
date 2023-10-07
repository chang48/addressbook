// Google Apps Script
// Convert multi-section Google Forms responses into columns

const responseSheetName = "Responses"
const splittedSheetName = "Splitted Responses"

function splitFormResponses() {
  const rs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(responseSheetName);
  const responses = rs.getRange(2, 1, rs.getLastRow() -1, rs.getLastColumn()).getValues();

  Logger.log(rs.getName())
  Logger.log(responseSheetName)
  Logger.log(splittedSheetName)

  const splittedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(splittedSheetName);
  const splittedResponses = [];
  var newResponse = [];
  var timestamp;
  var streetNumber;
  var city;
  var state;
  var zipCode;
  var photoURL;

  Logger.log('Response length = %s', responses.length)
  //Creating a 2D array using the amount of rows in the responses.
  for(i=0; i < responses.length; i++){
    //Loop through each cell in the row.
    for (j = 0; j < responses[i].length; j++) { 
      timestamp = responses[i][0];
      streetNumber = responses[i][1];
      city = responses[i][2];
      state = responses[i][3]
      zipCode = responses[i][4];
      photoURL = responses[i][5];

      //If there is "another" it will add the created row to the 2D array and restart the "newResponse" variable with the current row's timestamp and date
      if (responses[i][j] == "Yes") {

        splittedResponses.push(newResponse);
        newResponse = [timestamp, streetNumber, city, state, zipCode, photoURL];

      //If there is no "another" it will add the created row to the 2D array and jump to the next row in the original responses.
      } else if (responses[i][j] == "No"){

        splittedResponses.push(newResponse);
        newResponse = [];
        break;     

      //Otherwise, it will add the value of the current cell to the new row.
      } else { 
        newResponse.push(responses[i][j]);
      }
    }
  }

  //Printing the array.
  splittedSheet.getRange(2, 1, splittedResponses.length, splittedResponses[0].length).setValues(splittedResponses);
}
