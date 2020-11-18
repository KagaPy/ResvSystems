function pulldown(){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var row = sheet.getLastRow();
  
  var date =　sheet.getRange(row,6).getValue();
  
  sheet.getRange(row,12).setValue(date)
  
  var dateary = date.split(",")
  
  var rng = sheet.setActiveRange(sheet.getRange(row,6));
  
  Logger.log(dateary)
  
  var rule = 
      SpreadsheetApp
  .newDataValidation()
  .requireValueInList(dateary, true)
  .build();
  
  rng.setDataValidation(rule);
  
}