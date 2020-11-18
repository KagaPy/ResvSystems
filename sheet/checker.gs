function checker(e){
  //過去に参加してないか調べる
  //参加してたらすみませんが...メールを送る
  //使わない
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("joinedlist");

  var num_row = sheet.getLastRow(),
      participant = sheet.getRange(1,1,num_row,1).getValues();
  
  var joined = participant.some(function(array,index,data){
    var asheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
    var num_row = asheet.getLastRow();
    return(array[0] === asheet.getRange(num_row,3).getValue());
  })
  
  if(joined == true){
    
    var asheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答 1");
    
    var row = asheet.getLastRow(),
        name = asheet.getRange(row,3).getValue(),
        nmail = asheet.getRange(row,2).getValue();
    
    var thing = name +"様\n\n";
    thing += "この度は心理学実験へのご応募ありがとうございます。\n\n"
    thing += name + "様には既に同じ実験にご協力いただきましたため、申し訳ございませんが、予約をお取りできませんでした。"
    thing += "次の実験の際にも，ご協力いただけますと幸いです。\nどうぞよろしくお願いいたします。\n\nこころの未来研究センター";
    
    Logger.log(nmail);
    
    MailApp.sendEmail(nmail,"心理学実験：参加済み",thing, {name:"こころの未来研究センター"});
    
    //行削除
    asheet.deleteRow(row)
    
 }
}