// 移動させたいデータがあるシートのID
var ssId = "*************************************";

function main() {
  var ss = SpreadsheetApp.openById(ssId);
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var currentSS = SpreadsheetApp.getActiveSpreadsheet();
  var ssName = getSheetNames(ss);

  makeSheet(currentSheet, ssName);

  var data, dataSS;
  var data;
  for (var i = 0; i < ssName.length; i++) {
    dataSS = ss.getSheetByName(ssName[i]);
    data = getData(dataSS);
      
    data = replace0toEmpty(data);

    setData(currentSS, data, ssName[i]);
  }
}

function setData(currentSS, data, sheetName) {
  Logger.log(data);
 
  currentSS.getSheetByName(sheetName).getRange(1,1,data.length, data[0].length).setValues(data);
}


// 入力配列に沿ったシートを作成する
// 最初のデータのMetricsに沿ったシート作成をします！！
function makeSheet(currentSS, ssName) {

  var existSheetName = getSheetNames(currentSS);

  var isMatch;
  for(var i = 0; i < ssName.length; i++) {    
    isMatch = false;
    for (var j = 0; j < existSheetName.length; j++) {
      if(ssName[i] == existSheetName[j]) {
        currentSS.getSheetByName(ssName[i]).clear();
        isMatch = true;
        break;
      }
    }
    if(!isMatch) {
      currentSS.insertSheet(ssName[i]);
    }
  }
}


function replace0toEmpty(data) {

  var i = 0, j = 0, dataLength = data.length, data0Length = data[0].length;
  for(i = 0; i < dataLength; i++) {
    for (j = 0; j < data0Length; j++) {
      if(data[i][j] == 0) {
        data[i][j] = "";
      }
    }
  }

  return data;
}


function getData(ss) {
  var result = ss.getDataRange().getValues();
  result.shift();
  
  return result;
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