/* Main Method which orchestrates other methods */
function sendEmailmain() {
  const timestamp = new Date();
  logger(timestamp);
  //Button for Confirmation
  var confirm = Browser.msgBox(
    "Do you want to send Email?",
    Browser.Buttons.YES_NO
  );
  if (confirm != "yes") {
    logger("Exiting");
    return;
  } // if user click NO then exit the function, else the below is executed to send data
  sheet = getSheetData();
  patronList = getPatronList(sheet); //GetEmailList
  patronName = getPatronName(sheet); //GetPatronName List
  for (var i = 0; i < patronList.length; i++) {
    //For Each PatronEmail getPatronData and sendEmail
    //  if (patronList[i] == "anwarnishil@gmail.com") { //Used for testing
    logger(patronName[i]);
    patronData = getPatronData(sheet, patronList[i]);
    constructSendEmail(patronName[i], patronList[i], patronData);
    //  }
  }
  //
  //logger(patronData)
  Browser.msgBox("Script Completed successfully");
}

function logger(value) {
  //This is a logger function
  console.log(value);
}

function getSheetData() {
  //Returns Shee; Used for Reusing the Formated Data
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FormatedData");
  return sheet;
}
function constructSendEmail(patronName, patronEmail, patronData) {
  //Comments below
  //This function constructs html body and invokes sendemail function for each patrons
  var templateSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template"); //Get Templated Data
  sheetData = templateSheet.getDataRange().getValues();
  var subject = "";
  var bodyHeader = "";
  var bodyTail = "";
  var nextAMIA = "";
  //Get Templated Data from the cells.
  for (var i = 0; i < sheetData.length; i++) {
    if (sheetData[i][0] == "Subject") {
      subject = sheetData[i][1];
    }
    if (sheetData[i][0] == "BodyHeader") {
      bodyHeader = sheetData[i][1];
    }
    if (sheetData[i][0] == "BodyTail") {
      bodyTail = sheetData[i][1];
    }
    if (sheetData[i][0] == "NextAMIA") {
      nextAMIA = sheetData[i][1];
    }
  }
  bodyTail = bodyTail.replace("@", nextAMIA); //Replace BodyTail with nextAMIA date
  bodySalutation = "Assalamualaikum Dear " + patronName + ", <br/><br/>";
  table = createTable(patronData);
  bodyContent = `<html><head><style>table { border:1px solid #b3adad;
			border-collapse:collapse;
			padding:5px;
		}
		table th {
			border:1px solid #b3adad;
			padding:5px;
			background: #e87d88;
			color: #313030;
		}
		table td {
			border:1px solid #b3adad;
			text-align:center;
			padding:5px;
			background: #ffffff;
			color: #313030;
		}
	</style>
   </head><body><br>`;
  bodyContent = bodyContent + bodySalutation + bodyHeader + table + bodyTail;
  sendEmail(patronEmail, subject, bodyContent);
}
function sendEmail(patronEmail, subject, emailBody) {
  //SendEmail Function
  GmailApp.sendEmail(patronEmail, subject, "", { htmlBody: emailBody });
}
function createTable(patronData) {
  //This function creates the htmlTable for the email
  var table = "<table border=1>";
  var totalCols = 3; //Total only three Columsn
  for (var rowNo = 0; rowNo < patronData.length; rowNo++) {
    table = table + "<tr>"; //Table Row start
    if (rowNo == 0) {
      table = table + "<th>SlNo</th>"; //To Construct Serial Number
    } else {
      table = table + "<td>" + rowNo + "</td>"; //Table Serial Number
    }
    for (var colNo = 0; colNo < totalCols; colNo++) {
      colVal = patronData[rowNo][colNo];
      if (rowNo == 0) {
        table = table + "<th>" + colVal + "</th>"; //IF First Row Make it as Table header
      } else {
        table = table + "<td>" + colVal + "</td>"; //Columns
      }
    }
    table = table + "</tr>";
  }
  table = table + "</table>";
  return table;
}
function getPatronName(checklist) {
  //Gets PatronName from Formated Data
  var checklist_patronist = checklist.getRange("F2:F").getValues();
  var checklist_patrons_nums = checklist_patronist.filter(String).length;
  var arrName = [].concat.apply([], checklist_patronist).filter(String); //Removing Nulls by using a filter
  var result = [];
  for (var i = 0; i < arrName.length; i++) {
    result.push(arrName[i].split(" ")[0]);
  }
  return result;
}
function getPatronList(checklist) {
  //Gets PatrolEmailList from Formated Data
  var checklist_patronist = checklist.getRange("G2:G").getValues();
  var checklist_patrons_nums = checklist_patronist.filter(String).length; //Removing Nulls
  var result = [].concat.apply([], checklist_patronist).filter(String);
  return result;
}
function removeFilteredValues(range) {
  //Function used to remove filter from a sheet.
  const values = range.getValues();
  const firstRow = range.getRow();
  const sheet = range.getSheet();

  const filteredValues = values.filter((row, i) => {
    return !sheet.isRowHiddenByFilter(i + firstRow);
  });

  return filteredValues;
}
function getPatronData(sheet, patron) {
  //Get PatronData :-Details of Checkedout books
  var sheetRange = sheet.getRange("A1:E");
  var sheetData = sheet.getDataRange().getDisplayValues();
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
  sheetRange
    .createFilter()
    .setColumnFilterCriteria(
      2,
      SpreadsheetApp.newFilterCriteria().whenTextContains(patron).build()
    );
  //copy filtered data to temporary sheet
  var sourceRange = sheet.getFilter().getRange();
  const tempsheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Temporary");
  //delete sheet if it already exists
  if (tempsheet) {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempsheet);
  }
  //create temp sheet
  newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  newSheet.setName("Temporary");
  //copy to Temp Sheet
  sourceRange.copyTo(newSheet.getRange("A1"), { contentsOnly: true });
  var cell = newSheet.getRange("E1:E");
  cell.setNumberFormat("dd/mm/yyyy"); //Setting the Dateformat of CheckoutDate
  //cell.setValue(new Date()).setNumberFormat("dd/mm/yyyy");
  //SpreadsheetApp.getActiveSpreadsheet().deleteSheet(newSheet);
  newSheet.deleteColumns(1, 2); //Deleting first two columns Name and Email
  patronData = newSheet.getDataRange().getDisplayValues();
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(newSheet); //Delete temp sheet
  // clear our filter before leaving
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  return patronData;
}
