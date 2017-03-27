/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};
function onEdit(event){
  // sheet er hele regnearket
  var sheet = event.source.getActiveSheet();
  if(sheet.getName() == "Artikler") {  
    // rowedit udtrækker den række, der er blevet redigeret fra range-variablen i event. -1 for at kompensere for, at vi ikke bruger toprækken
    var rowedited = event.range.getRow() -1;
    // rangeall er hele det aktive område, der senere skal sorteres. Bemærk evt. at udvide hvis arket udvides
    var rangeall = sheet.getRange("A2:I99");
    // statuscell er cellen fra den første kolonne, hvor artiklens status står.
    var statuscell = rangeall.getCell(rowedited, 1);
    // statustext konverterer fra Object til en String
    var statustext = statuscell.getValue();
    statustext = statustext.toString();
    // sortindex er værdien i kolonne I, som vi bruger til at sortere efter
    var sortindex = rangeall.getCell(rowedited, 9);
      if(statustext == "kører") { sortindex.setValue(5);}
      if(statustext == "husk") { sortindex.setValue(4);}
      if(statustext == "standby") { sortindex.setValue(3);}
      if(statustext == "brød") { sortindex.setValue(2);}
      if(statustext == "idé") { sortindex.setValue(1);}
    // rækkerne med højeste værdi i kolonnen sortindex (kolonne I) sorteres øverst i arket, sekundært sorteres efter journalistnavn
    rangeall.sort([{column: 9, ascending: false}, {column: 3, ascending: true}]);
  }
  if(sheet.getName() == "Brød") {
    var rowedited = event.range.getRow();
    if(rowedited == 3){
      sheet.insertRows(7,1);
      var link = sheet.getRange("B2");
      link.copyTo(sheet.getRange("D7"));
      var rubrik = sheet.getRange("B3");
      rubrik.copyTo(sheet.getRange("B7"));
      var today = new Date();
      var datecell = sheet.getRange("C7");
      datecell.setValue(today);
      //var usercell = sheet.getRange("E7");
      //usercell.setValue(Session.getActiveUser().getEmail());
      var formatRow = sheet.getRange("A8:D8");
      formatRow.copyFormatToRange(sheet.getSheetId(), 1, 4, 7, 7);
      link.clearContent();
      rubrik.clearContent();
    }
    if(rowedited > 6){
      var rangeall = sheet.getRange("A7:I99");
      rowedited = rowedited - 6;
      var prioritycell = rangeall.getCell(rowedited, 1);
      var prioritytext = prioritycell.getValue();
      prioritytext = prioritytext.toString();
      var sortindex = rangeall.getCell(rowedited, 9);
      if(prioritytext == ""){ sortindex.setValue(1);}
      if(prioritytext == "lav"){ sortindex.setValue(2);}
      if(prioritytext == "høj"){ sortindex.setValue(5);}
      if(prioritytext == "top"){ sortindex.setValue(7);}
      rangeall.sort([{column: 9, ascending: false}, {column: 3, ascending: false}]);
    }
  }
}
function sortAll(){
  // opretter værdier i I-kolonnen til sortering. Se kommentarer ovenfor i onEdit
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var rangeopen = sheet.getRange("A2:I99");
  for (var j = 1; j <= 99; j++) {
    var statuscellopen = rangeopen.getCell(j, 1);
    var statuscelltext = statuscellopen.getValue();
    var statuscelltext = statuscelltext.toString();
    var sortindexopen = rangeopen.getCell(j, 9);
    if(statuscelltext == "") { sortindexopen.setValue(0);}
    if(statuscelltext == "kører") { sortindexopen.setValue(5);}
    if(statuscelltext == "husk") { sortindexopen.setValue(4);}
    if(statuscelltext == "standby") { sortindexopen.setValue(3);}
    if(statuscelltext == "brød") { sortindexopen.setValue(2);}
    if(statuscelltext == "idé") { sortindexopen.setValue(1);}
  }
  // sorterer arket
  rangeopen.sort([{column: 9, ascending: false}, {column: 3, ascending: false}]);
}
/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Sortér igen",
    functionName : "sortAll"
  }];
  /* der er ingen grund til at køre dette script hver gang. Kun hvis arket skal sorteres fra ny
  // opretter værdier i I-kolonnen til sortering, når arket åbnes. Se kommentarer ovenfor i onEdit
  var rangeopen = sheet.getRange("A2:I99");
  for (var j = 1; j <= 99; j++) {
    var statuscellopen = rangeopen.getCell(j, 1);
    var statuscelltext = statuscellopen.getValue();
    var statuscelltext = statuscelltext.toString();
    var sortindexopen = rangeopen.getCell(j, 9);
    if(statuscelltext == "") { sortindexopen.setValue(0);}
    if(statuscelltext == "kører") { sortindexopen.setValue(5);}
    if(statuscelltext == "opstart") { sortindexopen.setValue(4);}
    if(statuscelltext == "standby") { sortindexopen.setValue(3);}
    if(statuscelltext == "stoppet") { sortindexopen.setValue(2);}
    if(statuscelltext == "idé") { sortindexopen.setValue(1);}
  }
  // sorterer arket ved åbning
  rangeopen.sort([{column: 9, ascending: false}, {column: 3, ascending: false}]);  
  // scriptmenu fortsat fra originalt standardscript herunder
  */
  sheet.addMenu("Script Center Menu", entries);
};
