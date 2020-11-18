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
