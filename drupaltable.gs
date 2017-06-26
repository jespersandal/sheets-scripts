function createDrupalTable() {
  var ourSheet = SpreadsheetApp.getActiveSpreadsheet();
  var detteArk = ourSheet.getSheets()[0];
  var maxRows = detteArk.getLastRow();    
  var stringForOurRange = "A1:F" + maxRows;
  var ourRange = detteArk.getRange(stringForOurRange);
  detteArk.insertRows(2, 1);
  var overskriftCelle = ourRange.getCell(1, 6);
  overskriftCelle.setValue("Overskrift | Firma");
  var tabelCelle = ourRange.getCell(2, 6);
  tabelCelle.setValue("--|--");
  maxRows++;                                     // Vi har jo tilføjet en ekstra række ovenfor.  
  stringForOurRange = "A1:F" + maxRows;
  ourRange = detteArk.getRange(stringForOurRange);  
  for (var i = 3; i<= maxRows; i++) {
    var jobDescriptionCell = ourRange.getCell(i, 1);
    if (jobDescriptionCell.isBlank()) {
    }
    else {
      var jobDescription = jobDescriptionCell.getValue();
      jobDescription = jobDescription.toString();
      var jobLinkCell = ourRange.getCell(i, 5);
      var jobLink = jobLinkCell.getValue();
      jobLink = jobLink.toString();
      var jobCompanyCell = ourRange.getCell(i, 2);
      var jobCompany = jobCompanyCell.getValue();
      jobCompany = jobCompany.toString();
      var drupalString = "[" + jobDescription + "]" + "(" + jobLink + ")" + " | " + jobCompany;
      var outputCell = ourRange.getCell(i, 6);
      outputCell.setValue(drupalString); 
    }
  } 

}