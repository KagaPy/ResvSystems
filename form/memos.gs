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