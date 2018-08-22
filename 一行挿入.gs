// 挿入するデータの始まりのの日付
// 下記は2016/11/01 00:00:00.000を指定している
// ライブラリの挿入が必要ですよ
// Momentのライブラリkey = MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48
var firstDate = Moment.moment('2016/11/1');

function insertDateLine() {
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet();    
  var sheetNames = getSheetNames(currentSheet);

  // 仕事したのでコメントアウト
  // insertOneLine(currentSheet, sheetNames);
  
  var dateArr =  getDate(currentSheet, sheetNames);
  
  setDateArr(dateArr, currentSheet, sheetNames);
}

function setDateArr(dateArr, currentSheet, sheetNames) {

  var ss;
  
  for (var i = 0; i < sheetNames.length; i++) {
    ss = currentSheet.getSheetByName(sheetNames[i]);
    ss.getRange(2, 1).setValue("DATE");
    ss.getRange(3, 1, dateArr.length, 1).setValues(dateArr);  
  }
}



function insertOneLine(currentSheet, sheetNames) {
  var sheet;
  for (var i = 0; i < sheetNames.length; i++) {
     sheet = currentSheet.getSheetByName(sheetNames[i]);
     sheet.insertColumns(1, 1);
  }
}


function getDate(currentSheet, sheetNames) {
  var date = firstDate;
  var dateRow = [];
  var lastRow = currentSheet.getSheetByName(sheetNames[0]).getLastRow();
  
  var temp;
  for (var i = 0; i < lastRow; i++) {
    dateRow.push([date.clone().add(i, 'd').format('YYYY/M/D')]);
  }

  return dateRow;
}


// 既存のシート名一覧を取得する
function getSheetNames(currentSheet){
  var sheets = currentSheet.getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      sheet_names.push(sheets[i].getName());
    }
  }
  return sheet_names;
}