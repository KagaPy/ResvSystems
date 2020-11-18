function reminder(e){
 
 //しばらく動かさない
 var moment = Moment.load()
 
 var m = moment()
 
 //有効なGooglesプレッドシートを開く
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

 //最終行番号を取得
 var num_row = sheet.getLastRow();

 //列追加対策(2019/6/5)
 var cond = sheet.getRange(1,8).getValue();
  
  //選択肢更新で行がずれるの対策
  if(cond === "reminded"){

  for(var n=2; n < num_row+1;n++){
    
    var reminded = sheet.getRange(n,8).getValue(),
        nmail = sheet.getRange(n,2).getValue(),
        name = sheet.getRange(n,3).getValue(),
        date = sheet.getRange(n,6).getValue();
    
    var dstr = m.year() + date.slice(0,6);
    Logger.log(dstr);
    var ndate = moment(dstr, "YYYYMM/DD");
   
    var df = ndate.diff(m, "h")
    Logger.log(df);
    
    var thing_rmd = name+"様\n\n"
        + "お忙しい折失礼いたします。こちらはリマインドのメールですので、ご返信は不要です。\n\n"
        + "明日は実験へのご参加の方、よろしくお願いいたします。\n\n"
        +  name + "様の実験時間は、" + date + "です。\n"
        + "開始時間に遅れられる場合や，参加が不可能な場合は，当アドレスまでご連絡をお願いいたします。\n"
        + "連絡をいただけず15分遅れられた場合はキャンセル扱いとなりますのでご注意ください。\n"
        + "よろしくお願いいたします。\n\n"
        + "こころの未来研究センター"

    
    if(reminded == 1){
      
    }
    
    else if(df < 13){
      GmailApp.sendEmail(nmail,"心理学実験(リマインダ)",thing_rmd,{name:"こころの未来研究センター"});
      sheet.getRange(n,8).setValue(1);
      var list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("joinedlist");
      var lastrow = list.getLastRow();
      list.getRange(lastrow+1, 1).setValue(name);
      Logger.log("remind")
    }
    else{
   }
  }
  }
}

function setTriggerforrmd(){
  //12:30に↑を起動するトリガー
  var setTime = new Date();
  setTime.setHours(12);
  setTime.setMinutes(30); 
  ScriptApp.newTrigger('reminder').timeBased().at(setTime).create();
}