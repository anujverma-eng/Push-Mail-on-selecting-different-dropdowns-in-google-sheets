function addTODO(objectArray) {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Az75tKCyc3F08TuKaDHHQ-UQDX3V04oyD1_NW0ewoI8/edit#gid=0").getSheetByName(`${names.other7}`);
    ss.appendRow(objectArray);
  }
  
  function getTime() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timeZone = ss.getSpreadsheetTimeZone();
    var format = "dd/MM/yyyy HH:mm:ss";
    var output = Utilities.formatDate(new Date(), timeZone, format);
    Logger.log(output);
    return output;
  }
  
  function getLinkToRange(range) {
    var link = `${emails.other7}&range=${range}`
    return link;
  }
  