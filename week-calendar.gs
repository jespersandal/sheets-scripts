
/** This algorithm is a JavaScript port of the work presented by Claus Tøndering.
     
*/
function getWeek() {
       var a, b, c, d, e, f, g, n, s, w;
       var dagsDato = new Date();
       y = dagsDato.getFullYear();
       m = dagsDato.getMonth() + 1;
       d = dagsDato.getDate();

        if (m <= 2) {
            a = y - 1;
            b = (a / 4 | 0) - (a / 100 | 0) + (a / 400 | 0);
            c = ((a - 1) / 4 | 0) - ((a - 1) / 100 | 0) + ((a - 1) / 400 | 0);
            s = b - c;
            e = 0;
            f = d - 1 + (31 * (m - 1));
        } else {
            a = y;
            b = (a / 4 | 0) - (a / 100 | 0) + (a / 400 | 0);
            c = ((a - 1) / 4 | 0) - ((a - 1) / 100 | 0) + ((a - 1) / 400 | 0);
            s = b - c;
            e = s + 1;
            f = d + ((153 * (m - 3) + 2) / 5) + 58 + s;
        }
       
        g = (a + b) % 7;
        d = (f + g - e) % 7;
        n = (f + 3 - d) | 0;

        if (n < 0) {
            w = 53 - ((g - s) / 5 | 0);
        } else if (n > 364 + s) {
            w = 1;
        } else {
            w = (n / 7 | 0) + 1;
        }
       
        y = m = d = null;
       
        return w;
    };

function addWeek(){
  var regneark = SpreadsheetApp.getActiveSpreadsheet();
  var skabelonark = regneark.getActiveSheet();
  var aktueltarknavn = skabelonark.getSheetName();
  var opdeltstreng = aktueltarknavn.split(' ');
  var aktueltarkugenr = opdeltstreng[1]
  var nyuge = parseInt(aktueltarkugenr);
  if (nyuge >= 52) {
    nyuge = 1;
  }
  else {
    nyuge++;
  }
  var nytarknavn = "Uge " + nyuge.toString();
  regneark.insertSheet(nytarknavn, {template: skabelonark});
  var aktueltark = regneark.getActiveSheet();
  var aktueltarknavn = aktueltark.getSheetName();
  var aktueltarknavncelle = aktueltark.getRange("A1:A1");
  aktueltarknavncelle.clearContent();
  aktueltarknavncelle.setValue(aktueltarknavn);
  var cellertilrydning = aktueltark.getRange("B2:F12");
  cellertilrydning.clearContent();
}

function rydCeller(){
  var regneark = SpreadsheetApp.getActiveSpreadsheet();
  var cellertilrydning = regneark.getRange("B2:F12");
  cellertilrydning.clearContent();
  var aktueltark = regneark.getActiveSheet();
  var aktueltarknavn = aktueltark.getSheetName();
  var aktueltarknavncelle = aktueltark.getRange("A1:A1");
  aktueltarknavncelle.clearContent();
  aktueltarknavncelle.setValue(aktueltarknavn);
}

// Nedenstående er ikke færdigudviklet
/*
function udfyldDato(){
  var regneark = SpreadsheetApp.getActiveSpreadsheet();
  var aktueltfaneblad = regneark.getActiveSheet();
  var faneindex = aktueltfaneblad.getIndex();
  var forrigefane = (faneindex - 1);
  var sidstedato = regneark.getSheetValues(2, 6, 1, 1)[forrigefane];
  var celle = aktueltfaneblad.getRange("B2");
  
  var texttemp = sidstedato[1][1];
  texttemp = texttemp.toString();
  celle.setValue(texttemp);
  for (i = 0; i<=4; i++){
    var 
    
  
  
  
} 
*/

//Nedenstående løser måske onOpen-problemet. Det er en kopi af onOpen med sin egen trigger.

function openCurrentWeek() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();    
  var ugeNummer = getWeek();
  sheet.setActiveSheet(sheet.getSheetByName("Uge "+ugeNummer));
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tilføj og udfyld")
    .addItem("Gå til aktuel uge", "openCurrentWeek")
    .addItem("Tilføj ny uge", "addWeek")
    .addItem("Ryd vagtplanen", "rydCeller")
    .addToUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ugeNummer = getWeek();
  sheet.setActiveSheet(sheet.getSheetByName("Uge "+ugeNummer));
};
