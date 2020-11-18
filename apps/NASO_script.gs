
//DBID = "17vxz_Bohs5O3gFMPnhf71oSCh-8mdD3jouIkZ-qmwlM"
//moment.js = MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48


function setting_1(){
  layoutSettings();
}

function setting_2(){
  triggerSettings();
}

//ーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー

function layoutSettings(){

  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      formAns = ss.getSheetByName("フォームの回答 1");

  formAns.getRange(1,7).setValue("進捗")

  var dateary = ["保留","調整","確定"],
      rng = formAns.setActiveRange(formAns.getRange(2,7,formAns.getMaxRows(),1)),
      rule = SpreadsheetApp.newDataValidation().requireValueInList(dateary, true).build();

  rng.setDataValidation(rule)

  formAns.getRange(1,8).setValue("リマインダ")
  formAns.getRange(1,9).setValue("備考")
  formAns.getRange(1,10).setValue("入力日")
  formAns.setFrozenRows(1);
  formAns.getRange(1,1,1,formAns.getLastColumn()).createFilter();
  formAns.deleteColumns(formAns.getLastColumn()+1, formAns.getMaxColumns()-formAns.getLastColumn())

  ss.insertSheet("実験概要",2)

  var meta = ss.getSheetByName("実験概要");
  var title = ss.getName().slice(0,-4);

  var db = SpreadsheetApp.openById("17vxz_Bohs5O3gFMPnhf71oSCh-8mdD3jouIkZ-qmwlM"),
      expdb = db.getSheetByName("実験DB"),
      roomdb = db.getSheetByName("実験ルームDB"),
      workerdb = db.getSheetByName("実験者db");

  var enames = expdb.getRange(3,1,expdb.getLastRow()-2,1).getValues();
  var list = []

  for(var i=0;i<enames.length;i++){

    if(enames[i][0]!=title){}
    else{
      for(var j=1;j<14;j++){
        var v = [expdb.getRange(i+3,j).getValue()];
        list.push(v);
      }
    }
  }

  var tag = ["実験名","担当","実験者","所属","重複不可","場所","謝礼","人数","所要","間隔","昼休","フォームURL","フォームID","実験室calID"]
  title = roomdb.getLastRow()

  var roomID = roomdb.getRange(3,2,roomdb.getLastRow()-2).getValues(),
      roomID = Array.prototype.concat.apply([],roomID),
      roomID = roomID.indexOf(list[5][0])+1,
        roomID = roomdb.getRange(roomID+2,3).getValue();

  list.push(roomID)

  var works = list[2][0].split(",")
  Logger.log(works+":"+typeof(works)+works.length)

  for(i=0;i<works.length;i++){

    var workID = workerdb.getRange(3,3,workerdb.getLastRow()-2).getValues(),
        workID = Array.prototype.concat.apply([],workID),
        workID = workID.indexOf(works[i])+1;

    var wmail = workerdb.getRange(workID+2,4).getValue();
    var workID = workerdb.getRange(workID+2,5).getValue();

    tag.push(works[i])
    list.push(workID);
    list.push(wmail);

  }

  for(i=0;i<tag.length;i++){
    if(i<14){
      meta.getRange(i+1,1).setValue(tag[i]);
      meta.getRange(i+1,2).setValue(list[i]);
    }else{
      meta.getRange(i+1,1).setValue(tag[i]);
    }
  }

  //カレンダーIDとOAemail
  for(i=14;i<list.length;i=i+2){
    var row = meta.getRange(1,2,meta.getLastRow()).getValues(),
        row = row.filter(String).length;
    meta.getRange(row+1,2).setValue(list[i])
    meta.getRange(row+1,3).setValue(list[i+1])
  }
  //1.3
  title = Array.prototype.concat.apply([],workerdb.getRange(1,3,workerdb.getLastRow(),1).getValues()).indexOf(list[1][0])+1
  meta.getRange(2,3).setValue(workerdb.getRange(title,4).getValue())
  meta.getRange(meta.getLastRow()+1,1).setValue("備考") //最後-1
  meta.getRange(meta.getLastRow()+1,1).setValue("募集状態") //最後

}

function triggerSettings(){

  ScriptApp.newTrigger("setRmd")
  .timeBased()
  .atHour(0)
  .everyDays(1)
  .create();

  ScriptApp.newTrigger("reservationMaker")
  .timeBased()
  .everyMinutes(15)
  .create();

  var sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("applicationChecker")
  .forSpreadsheet(sheet)
  .onFormSubmit()
  .create();

  ScriptApp.newTrigger("selectDates")
  .timeBased()
  .everyHours(1)
  .create();

  selectDates()

}

//----------------------------------------------------------------------------------------------------------------

//SSから引数を引っ張ってきていれる
function reservationMaker(e){

  var moment = Moment.load(),
      m = moment();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("実験概要");

  var exp_name = sheet.getRange(1,2).getValue(),
      belonging = sheet.getRange(4,2).getValue(),
      r_name = sheet.getRange(2,2).getValue(),
      place = sheet.getRange(6,2).getValue(),
      reward = sheet.getRange(7,2).getValue(),
      num = sheet.getRange(8,2).getValue(),
      roomID = sheet.getRange(14,2).getValue(),
      calID = sheet.getRange(15,2).getValue(),
      bikou = sheet.getRange(sheet.getLastRow()-1,2).getValue();

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答 1");
  var row = sheet.getLastRow();

  Logger.log(roomID)

  for(var i=2; i<row+1; i++){

    var process = sheet.getRange(i,7).getValue(),
        rmd = sheet.getRange(i,8).getValue();

    if(process == "調整"){

      var name = sheet.getRange(i,3).getValue(),
          date = sheet.getRange(i,6).getValue(),
          age = sheet.getRange(i,4).getValue(),
          sex = sheet.getRange(i,5).getValue();

      var thing =
          name + "様\n\n"
          + "心理学実験にご応募頂きありがとうございます。実験の予約を受け付けました。以下の内容のご確認をお願いいたします。\n\n"
          + "【実験名】"+ exp_name + "\n\n"
          + "【日時】"+ date +"\n\n"
          + "【謝礼】"+ reward +"\n\n"
          + "【場所】"+place+"\n\n"
          + "【備考】"+bikou+"\n\n"
          + "日時等の修正・ご質問等ございましたらお気軽にお問い合わせ下さい。\n\n"
          + "どうぞよろしくお願いいたします。\n\n"
          + belonging +" "+ r_name;

      var smom = m.year() + date.slice(0,6) + date.slice(12,19);
      var emom = m.year() + date.slice(0,6) + date.slice(22,29);

      var sdate = moment(smom, "YYYYMM/DD HH:mma").toDate();
      var edate = moment(emom, "YYYYMM/DD HH:mma").toDate();

      var room = CalendarApp.getCalendarById(roomID);
      Logger.log(room)
      //1人で使用することを想定
      var shift = CalendarApp.getCalendarById(calID);

      var yoyaku = room.getEvents(sdate, edate);

      if(yoyaku.length > num-1){

        sheet.getRange(i,7).setValue("実験室は予約済みです！");

      }else if(shift.getEvents(sdate, edate) == 0){

        sheet.getRange(i,7).setValue("勤務予定がありません！");

      }else if(date.length == 0){

        sheet.getRange(i,7).setValue("日程未選択です！");

      }else if(yoyaku.length<num+1 && rmd != 1){

        var oa = shift.getEvents(sdate, edate),
            email = sheet.getRange(i,2).getValue();

        var put = r_name +"("+name+"_"+sex+"_"+ age +")";

        room.createEvent(put, sdate, edate)
        shift = shift.createEvent(put, sdate, edate)
        shift.addEmailReminder(720
                               + moment(smom,"YYYYMM/DD HH:mma").diff(moment(smom,"YYYYMM/DD 00:00"),"m"))

        //怒られ発生を防ぐ
        Utilities.sleep(5000)

        GmailApp.sendEmail(email,"心理学実験",thing,{name:belonging});

        sheet.getRange(i,7).setValue("確定");

      }
    }
  }
}


//－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－


function Reminder(e){

  var moment = Moment.load()

  var m = moment()

  //有効なGooglesプレッドシートを開く
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = ss.getSheetByName("実験概要");

  var exp_name = sheet.getRange(1,2).getValue(),
      belonging = sheet.getRange(4,2).getValue(),
      r_name = sheet.getRange(2,2).getValue(),
      place = sheet.getRange(6,2).getValue(),
      r_email = sheet.getRange(2,3).getValue();

  sheet = ss.getSheetByName("フォームの回答 1")

  //リマインダ列番号
  var rmd = getCol("リマインダ",sheet), //確実に
      row = sheet.getLastRow(),
      cond = sheet.getRange(1,rmd).getValue();

  //選択肢更新で行がずれるの対策
  if(cond === "リマインダ"){

    var participant = [];

    for(var n=2; n<row+1;n++){

      var check = sheet.getRange(n,rmd),
          nmail = sheet.getRange(n,2).getValue(),
          name = sheet.getRange(n,3).getValue(),
          date = sheet.getRange(n,6).getValue();

      var dstr = m.year() + date.slice(0,6);
      var ndate = moment(dstr, "YYYYMM/DD");

      var df = ndate.diff(m, "h")

      var thing_rmd = name+"様\n\n"
      + "お忙しい所失礼いたします。こちらはリマインドのメールですので、ご返信は不要です。\n\n"
      + "明日は実験("+exp_name+")へのご参加をよろしくお願いいたします。\n\n"
      +  name + "様の実験時間は、" + date + "です。\n\n"
      + "実験場所は"+place+"です。\n\n"
      + "参加が不可能になられた場合には，ご連絡をお願いいたします。\n\n"
      + "明日はよろしくお願いいたします。\n\n"
      + belonging +" "+ r_name

      var go = df>0 && df<13? true:false; //時間

      if(check.getValue() == 1){}

      else if(sheet.getRange(n,7).getValue() === "確定" && go === true){

        GmailApp.sendEmail(nmail,"心理学実験(リマインダ)", thing_rmd, {name:belonging,bcc:r_email});
        check.setValue(1);
        name = name.replace(" ","").replace("　","")
        participant.push(name)

      }else{}
    }

    ss = SpreadsheetApp.openById("17vxz_Bohs5O3gFMPnhf71oSCh-8mdD3jouIkZ-qmwlM");
    sheet = ss.getSheetByName("参加者DB");

    var index = (sheet.getRange(2,1,1,sheet.getLastColumn()).getValues()[0].indexOf(exp_name))+1;
    row = sheet.getRange(1,index,sheet.getLastRow(),1).getValues();
    row = row.filter(String).length;

    //ここは問題

    for(n=0;n<participant.length;n++){
      sheet.getRange(row+1,index).setValue(participant[n]);
      row = row+1
    }

  }

}

//－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

function setRmd(){
  var setTime = new Date();
  setTime.setHours(12);
  setTime.setMinutes(30);
  ScriptApp.newTrigger("Reminder").timeBased().at(setTime).create();
}

//列名を引数として列番号を返す関数
function getCol(string, sheet) {

  var lastCol = sheet.getLastColumn(),
      range = sheet.getRange(1,1,1,lastCol),
      values = range.getValues()[0];

  for (i=0; i<lastCol; i++) {
    if (values[i] == string) {
      return i + 1; // カラム番号を返す
    }
  }
  return false; // 存在しなければfalse
}

//－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

function applicationChecker(){
  //過去に参加してないか調べる
  //参加してたらすみませんが...メールを送る
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答 1"),
      meta = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("実験概要"),
      rest = meta.getRange(5,2).getValue().split(",")

  var row = sheet.getLastRow(),
      name = sheet.getRange(row,3).getValue();

  var partDB = SpreadsheetApp.openById("17vxz_Bohs5O3gFMPnhf71oSCh-8mdD3jouIkZ-qmwlM").getSheetByName("参加者DB")

  var check = 1

  for(var i=0;i<rest.length;i++){

    var index = partDB.getRange(2,1,1,sheet.getLastColumn()).getValues(),
        index = Array.prototype.concat.apply([],index),
        index = (index.indexOf(rest[i]))+1;

    var part = partDB.getRange(1,index,partDB.getLastRow(),1).getValues(),
        npart = part.filter(String).length,
        part = partDB.getRange(1,index,npart,1).getValues(),
        part = Array.prototype.concat.apply([],part);

    var str = name.replace(" ","").replace("　","")

    var check = part.indexOf(str)<0? check*1:check*0;
  }

  //
  if(check<1){

    var belonging = meta.getRange(4,2).getValue(),
        r_name = meta.getRange(2,2).getValue(),
        nmail = sheet.getRange(row,2).getValue();

    var thing =
        name +"様\n\n"
        + "この度は心理学実験へのご応募ありがとうございます。\n\n"
        + "申し訳ございませんが、" + name + "様には既に同様の実験にご協力頂きましたため、予約をご用意できませんでした。"
        + "次の実験の際にも，ご協力いただけますと幸いです。\n\n"
        + "どうぞよろしくお願いいたします。\n\n"
        + belonging+" "+r_name;

    MailApp.sendEmail(nmail,"心理学実験",thing,{name:belonging});

    sheet.deleteRow(row)

    return

  }

  if(row<2){return}

  var col = sheet.getLastColumn(),
      date =　sheet.getRange(row,6).getValue();

  sheet.getRange(row,col).setValue(date)

  var dateary = date.split(","),
      rng = sheet.setActiveRange(sheet.getRange(row,6)),
      rule = SpreadsheetApp.newDataValidation().requireValueInList(dateary, true).build();

  rng.setDataValidation(rule);

  thing = "名前："+name+"\n\n"+"日時："+date;

  MailApp.sendEmail(meta.getRange(2,3).getValue(),meta.getRange(1,2).getValue()+(row-1)+"人目",thing)

}


function selectDates(){

  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      meta = ss.getSheetByName("実験概要");

  var form = FormApp.openById(meta.getRange(13,2).getValue());

  var accept = form.isAcceptingResponses()

  if(accept === true){

    form.setAcceptingResponses(false);
    meta.getRange(meta.getLastRow(),2).setValue("停止中")
    form.setCustomClosedFormMessage("現在選択肢の更新中です。\n数分後に再度のアクセスをお願い致します。");

    var item = form.getItems(FormApp.ItemType.CHECKBOX),
        sheet = ss.getSheetByName("フォームの回答 1");

    var day = Utilities.formatDate(new Date(), "JST", "(M/d/hh:mm)");

    var num_row = sheet.getLastRow();
    var check = sheet.getRange(1,7).getValue();

    if(num_row >1){
      var record = sheet.getRange(2,6,num_row-1,1).getValues(),
          record = Array.prototype.concat.apply([],record);

      var process = sheet.getRange(2,7,num_row-1,1).getValues();
      var go = every(process,"確定");
    }else{
      var record = sheet.getRange(2,6,1,1).getValues(),
          record = Array.prototype.concat.apply([],record);
      var go = true
      }

    if(check == "進捗" && item != null && go == true){

      var newchoices = calendarDiv(meta),
          num = meta.getRange(8,2).getValue();

      for(var m=0;m<record.length;m++){

        record[m] = record[m].slice(0,1) == " "? record[m].slice(1,30):record[m];
        var cnt = 0

        for(var n=0;n<newchoices.length;n++){
          cnt = newchoices[n] === record[m]? cnt+1:cnt;
        }

        var times = cnt >= num+1? cnt:cnt-1 ;

        var tg = newchoices.indexOf(record[m])

        for(var a=0;a<times;a++){
          newchoices.splice(tg,1)
          tg = newchoices.indexOf(record[m])
        }

      }

      if(newchoices != 0){

        var chbox = form.addCheckboxItem().setTitle("希望実験日時").setChoiceValues(newchoices);
        var valid = meta.getRange(9,2).getValue(),
            serge = valid < 60? 5:(valid < 90? 3:(valid<120? 2:1));

        serge = newchoices.length<=serge-3? 1:serge;

        valid = FormApp.createCheckboxValidation()
        .setHelpText("少なくとも"+serge+"つ選択してください")
        .requireSelectAtLeast(serge)
        .build();

        chbox.setRequired(true)
        chbox.setValidation(valid)

      }
      else{
        //調整済みのニューアイテムを作成
        form.addCheckboxItem().setTitle("希望実験日時").setChoiceValues(["no time"])
      }
      //以下の処理は6秒遅延（そうしないとリマインダのところに予約時間が入ってしまうため）
      Utilities.sleep(5000)

      //新列にrecordの結果を格納
      for(var i=0;i<record.length;i++){
        sheet.getRange(i+2,7).setValue(record[i]);
      }

      form.deleteItem(3)
      Utilities.sleep(10000)
      sheet.deleteColumn(6)

    }else{
      return
    }
    form.setAcceptingResponses(true);
    meta.getRange(meta.getLastRow(),2).setValue("募集中");
  }
  else{
    return
  }

}

//－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－


function calendarDiv(meta){

  var moment = Moment.load(),
      calID = meta.getRange(15,2).getValue(),
      span_m = meta.getRange(9,2).getValue(),
      worker = meta.getRange(3,2).getValue();
  Logger.log(worker)

  //当日を取得
  var startDate = new Date(); //取得開始日
  startDate.setDate(startDate.getDate()+1);
  var endDate = new Date();
  endDate.setDate(endDate.getDate()+10);

  var oas = CalendarApp.getCalendarById(calID).getEvents(startDate, endDate,{search:worker});

  var list = [];

  for(var j=0;j<oas.length;j++){
    var Start = moment(oas[j].getStartTime()); //イベントの開始時刻
    var End= moment(oas[j].getEndTime()); //イベントの終了時刻
    var koma = End.diff(Start, "m")/span_m < 1? 0:Math.floor(End.diff(Start, "m")/span_m);
    Logger.log(koma)
    //koma = span_m * koma <//コマ数の取得(母数はコマの分の長さ)

    var sdate = oas[j].getStartTime()
    var komali = []
    //Logger.log(sdate)

    for(var i=0;i<koma;i++){
      var skoma = new Date(sdate.getFullYear(),sdate.getMonth(),sdate.getDate(),sdate.getHours(),sdate.getMinutes()+span_m*(i),0);
      var ekoma = new Date(sdate.getFullYear(),sdate.getMonth(),sdate.getDate(),sdate.getHours(),sdate.getMinutes()+span_m*(i+1),0);
      var select = moment(skoma).format('MM/DD (ddd) hh:mma')+" - "+moment(ekoma).format('hh:mma')
      komali.push(select);
    }

    list.push(komali);

  }

  list = Array.prototype.concat.apply([],list)

  return list

}


//－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－


function getCol(string, sheet){


  var lastCol = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, 1, lastCol)
  var values = range.getValues()[0];

  for (i=0; i<lastCol; i++) {
    if (values[i] == string) {
      return i + 1; // カラム番号を返す
    }
  }
  return false; // 存在しなければfalse
}


//－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－


function every(array,value){

  var contents = Array.prototype.concat.apply([],array);

  var counts = 0

  for(var n=0;n<contents.length;n++){

    if(contents[n] == value){
      counts = counts +1 }
    else{

    }

  }

  return(counts == contents.length);

}
