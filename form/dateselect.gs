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