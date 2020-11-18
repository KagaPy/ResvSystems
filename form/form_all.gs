function dateselect(e){
  
  var form = FormApp.getActiveForm();
  
  var item = form.getItems(FormApp.ItemType.CHECKBOX);
  
  //これまでの日時回答の格納
  var ssid = "1YuJMbS76pWOOfw0UkNBFlk3rbGXV0IzYzg8CyDNEQ-c";
  var ss = SpreadsheetApp.openById(ssid).getSheetByName("フォームの回答 1");
  var num_row = ss.getLastRow()
  //Logger.log(num_row)
  
  var rec = ss.getRange(2, 6, num_row-1, 1).getValues();
  var record = Array.prototype.concat.apply([],rec);
  var long = record.length;
  //Logger.log(record+long);
  
  //全部確定かどうかチェック
  var namae = ss.getRange(1,7).getValue();
  Logger.log(namae)
  var process = ss.getRange(2,7,num_row-1,1).getValues();
  var go = every(process,"確定")
  
  if(namae == "進捗" && item != null && go == true){
      
      //アイテム構成
      //新たな日時の用意
      var newchoices = Calendardiv("7k4v88umnk9in3qsov3cugkovg@group.calendar.google.com", 30);
      Logger.log(newchoices);
      
      //recordに入ってるものについては消す
      for(var m=0;m<long;m++){
        var data = record[m]
        var tg = newchoices.indexOf(data)
        Logger.log(tg)
        //2回やっとかないと予約が入ったことで2回取り出されたものが消えない
        if(tg > -1.0){
          newchoices.splice(tg,1)
          var tg2 = newchoices.indexOf(data)
          
          if(tg2 > -1.0){
            newchoices.splice(tg2,1)
          } 
        }
      }
      
      if(newchoices != 0){
        //調整済みのニューアイテムを作成
        form.addCheckboxItem().setTitle("希望実験日時").setChoiceValues(newchoices)
        
        //以下の処理は6秒遅延（そうしないとリマインダのところに予約時間が入ってしまうため）
        Utilities.sleep(6000)
        
        //新列にrecordの結果を格納
        for(var i=0;i<long;i++){
          ss.getRange(i+2,7,1,1).setValue(record[i]);
        }
        
        //アイテムの削除
        form.deleteItem(3)
        
        //
        Utilities.sleep(5000)
        
        //旧列の削除
        //なぜか6列目がある扱いされるときとない扱いされるときがある
        //???
        ss.deleteColumns(6)
      }
    }
}


function Calendardiv(calid,span){
 //シフトのカレンダーにある時間をコマ数分割してカレンダーに記載可能なDateオブジェクトとフォームに表示する文字列を作成する
 var moment = Moment.load() 
 
 var oawork = CalendarApp.getCalendarById(calid);
　//当日を取得
 var startDate = new Date(); //取得開始日
 startDate.setDate(startDate.getDate()+1);
 var endDate = new Date();
 endDate.setDate(endDate.getDate()+20);　//取得終了日は20日先
 
  var oas = oawork.getEvents(startDate, endDate);
 //Logger.log(oas)

 var list = [] 
 
 for(var j=0;j<oas.length;j++){
    var Start = moment(oas[j].getStartTime()); //イベントの開始時刻
    var End= moment(oas[j].getEndTime()); //イベントの終了時刻
    var koma = End.diff(Start, "m")/span //コマ数の取得(母数はコマの分の長さ)
    
    var sdate = oas[j].getStartTime()
    var komali = []
    //Logger.log(sdate)
    
    for(var i=0;i<koma;i++){
      var skoma = new Date(sdate.getFullYear(),sdate.getMonth(),sdate.getDate(),sdate.getHours(),sdate.getMinutes()+span*(i),0);
      var ekoma = new Date(sdate.getFullYear(),sdate.getMonth(),sdate.getDate(),sdate.getHours(),sdate.getMinutes()+span*(i+1),0);
      var select =moment(skoma).format('MM/DD hh:mma')+" - "+moment(ekoma).format('hh:mma')
      komali.push(select);
    }
   //Logger.log(komali)
   list.push(komali);
  }
  
  list = Array.prototype.concat.apply([],list)
  
  return list
}


//列名を引数として列番号を返す関数
function getCol(string, id) {

  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getActiveSheet();
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

function test(){
  var result = getCol("希望実験日時", "1YuJMbS76pWOOfw0UkNBFlk3rbGXV0IzYzg8CyDNEQ-c")
  Logger.log(result)
}

//配列内のすべての値が特定の文字列に一致してたらtrueを返す
//普通everyとしてありそうだけど
function every(array,value){
  
  var contents = Array.prototype.concat.apply([],array);
  
  var counts = 0
  
  for(var n=0;n<contents.length;n++){
    if(contents[n] == value){
    counts = counts +1
    }else{
    }
  }
  
  return(counts == contents.length);
  
}