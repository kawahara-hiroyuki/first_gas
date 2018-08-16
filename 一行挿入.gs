// 挿入するデータの始まりのの日付
// 下記は2016/11/01 00:00:00.000を指定している
var firstDate = new Date(2016, 10, 1, 0, 0, 00);

function insertDateLine() {
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet();    
  var sheetNames = getSheetNames(currentSheet);

  // 仕事したのでコメントアウト
  // insertOneLine(currentSheet, sheetNames);
  
  var dateArr =  getDate(currentSheet, sheetNames);
  
//  setDateArr(dateArr, currentSheet, sheetNames);
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

  for (var i = 0; i < lastRow; i++) {   
    dateRow.push([date.setDate(date.getDate() + 1).toString()]);  

    if(!isNaN(dateRow[i][0])) {
      dateRow[i][0] = new Date((dateRow[i][0] - 25569) * 86400000);
    }
  }
  
  Logger.log(dateRow);
  
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