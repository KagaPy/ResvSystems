function adjuster(e){
  //誰か送信したら知らせる
　var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var num_row = sheet.getLastRow();
  
  var name =　sheet.getRange(num_row,3).getValue(),
      choices = sheet.getRange(num_row,6).getValue(),
      thing = "名前：" + name + "\n"
  　　　　　　+ "日時：" + choices
  
  MailApp.sendEmail("psyexp20ns@gmail.com",num_row,name);
  
  Logger.log(name)

}