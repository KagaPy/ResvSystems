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


function MakeResv(e){
  
  var moment = Moment.load()
  
  var m = moment()
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答 1");
  
  var row = sheet.getLastRow();
  
  for(var i=2; i<row+1; i++){
    
    var check = sheet.getRange(i,7).getValue();
    
    var decide = sheet.getRange(i,8).getValue();
    
    Logger.log(check)
    
    if(check == "調整"){
      
      var name = sheet.getRange(i, 3).getValue();
      
      var date = sheet.getRange(i, 6).getValue();
      Logger.log(date);
      
      var thing =
          name + "様\n\n"
          + "心理学実験の予約を受け付けました。以下の内容をご確認ください。\n\n"
          + "【実験名】認知的柔軟性\n"
          + "【日時】"+date+"\n"
          + "【謝礼】Amazonギフト券 500円分(毎週水曜午前発送)\n"
          + "【場所】こころの未来研究センター別館(関田南研究棟内) 2F実験室\n"
          + "(アクセス http://kokoro.kyoto-u.ac.jp/jp/about/access.html"
          + " 〒606-8203　京都市左京区田中関田町2-24　※川端通り沿いの「本館」ではないのでご注意ください）\n\n"
          + "日時の修正が必要となられた場合はご返信ください。\n"
          + "当日はよろしくお願いいたします。\n\nこころの未来研究センター";
      
      Logger.log(thing);
      
      var smom = m.year() + date.slice(0,14);
      var emom = m.year() + date.slice(0,6) + "" + date.slice(16,24);
      Logger.log(smom);
      Logger.log(emom);
      
      var sdate = moment(smom, "YYYYMM/DD HH:mma").toDate();
      var edate = moment(emom, "YYYYMM/DD HH:mma").toDate();
      Logger.log(sdate)
      
      var room = CalendarApp.getCalendarById("di3kmmal7iki7ckrlbkt7bbq7k@group.calendar.google.com");
      
      var shift = CalendarApp.getCalendarById("7k4v88umnk9in3qsov3cugkovg@group.calendar.google.com");
      
      // 実験室に予約が無く，OA勤務可能が0でなければ続行
      if(room.getEvents(sdate, edate) != 0){
        
        sheet.getRange(i,7).setValue("実験室は予約済みです！");
        
      }else if(shift.getEvents(sdate, edate) == 0){
      
        sheet.getRange(i,7).setValue("勤務可能なOAがいません！");
      
      }else if(room.getEvents(sdate, edate) == 0){
        
        var oa = shift.getEvents(sdate, edate),
            email = sheet.getRange(i, 2).getValue();
        
        var put = "Ueda"+"("+name+")";

        //まだOAへのリマインド機能は実装していない
        
        room.createEvent(put, sdate, edate);
        shift.createEvent(put, sdate, edate);
        
        //怒られ発生を防ぐ
        Utilities.sleep(5000)
        
        GmailApp.sendEmail(email,"心理学実験",thing,{name:"こころの未来研究センター"});
        
        sheet.getRange(i,7).setValue("確定");
        
      }else if(check != 1){
      }
    }
  }
}
