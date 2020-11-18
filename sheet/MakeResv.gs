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