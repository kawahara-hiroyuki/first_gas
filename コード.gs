/*
残タスク
列数が26以上 A, B, ~ Z ,AAの場合の処理追加
キー入力項目をどうやって作ろうかな？
リファクタリング（大事）
date欲しくなった
グラフに落とすために、いらない数値もあったなぁ
*/



function main() {
  // GAの結果が詰まってるスプレッドシートのキーを入力させる
  // var key = getSpreadSheetKey();
  var key = "1UnoTHCWHHoPW_5Ryy8Jbxo0uW-TP6Cnwz8NiLJQCaog"; // 仮置き

  // シートの情報取ってくる系
  try{
    var ss = SpreadsheetApp.openById(key);
  }catch(e){
    Browser.msgBox(e);
  }
  
  // 取ってきたデータを入力する系
  setSheetData(ss, key);
}


// シートの情報を入力する系
function setSheetData(ss, key) {
  var reportConfiguration = ss.getSheetByName("Report Configuration").getDataRange().getValues();
  var reportNameRow = 2, metricsRow = 6, dimensionsRow = 7;
  
  var currentSheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 入力に必要なシートを準備する
  makeSpreadSheet(reportConfiguration[metricsRow - 1], currentSheet);

  // データのヘッダーを入力する
  // 拡張して、各データを入力するまで広げていく予定
  setHeader(currentSheet, reportConfiguration[reportNameRow - 1], reportConfiguration[metricsRow - 1]);

  // データを入力する
  setGAData(currentSheet, ss, reportConfiguration[reportNameRow - 1], reportConfiguration[metricsRow - 1], key);
}


function setGAData(currentSheet, ss, reportName, reportMetrics, key) {
  var metrics = reportMetrics[1].split(/\r\n|\r|\n/);        
  var dataRange = getDataRange(ss, reportName[1]);
  
  var queryArr = [];
  var aveQuery = [];
  var dataQuery = [];
  for(var i = 0; i < metrics.length; i++) {
    // 空行チェック
    if(metrics[i] == "") continue;
    
//    旧バージョン。よしなにしてくれるセット
//    queryArr = getQueryArr(key, dataRange, reportName.length ,i);
//    setQuery(currentSheet, queryArr, reportMetrics, i);

    // 新バージョン。　平均！シート名！データ！って感じにしてくれる
    aveQuery =  getAveQuery (key, dataRange, reportName.length ,i);
    setAveQuery(currentSheet, aveQuery, reportMetrics, i)
    
    dataQuery = getDataQuery(key, dataRange, reportName.length ,i)
    setDataQuery(currentSheet, dataQuery, reportMetrics, i);
        
  } 
}

function getAveQuery(key, dataRange, sheetLength, times) {
  var result = [];
  var string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var sheetRow = '';
  var dataRow = '';
  var query = '';

  // 平均値を出すようのクエリを追加
  // =IMPORTRANGE("1UnoTHCWHHoPW_5Ryy8Jbxo0uW-TP6Cnwz8NiLJQCaog", B2&"!G12:G12")を作成し、resultへpushする
  for (var i = 0; i < sheetLength; i++) {
    sheetRow = string.slice(i, i + 1);
    dataRow = string.slice(dataRange[1] + times - 1, dataRange[1] + times);
    query = '=IMPORTRANGE("' + key + '", ' + sheetRow +'2&"!' + dataRow + (dataRange[0] + 1) + ':' + dataRow + (dataRange[0] + 1) + '")';
    result.push(query);
  }
  
  return result;
}

function setAveQuery(currentSheet, queryArr, reportMetrics, times) {
  var metrics = reportMetrics[1].split(/\r\n|\r|\n/);

  var setQueryArr = [];
  setQueryArr.push(queryArr);

  currentSheet.getSheetByName(metrics[times]).getRange(1, 1, 1, queryArr.length).setValues(setQueryArr);
}

function getDataQuery(key, dataRange, sheetLength, times) {
  var result = [];
  var string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var sheetRow = '';
  var dataRow = '';
  var query = '';

  // 平均値を出すようのクエリを追加
  // =IMPORTRANGE("1UnoTHCWHHoPW_5Ryy8Jbxo0uW-TP6Cnwz8NiLJQCaog", B2&"!G12:G12")を作成し、resultへpushする
  for (var i = 0; i < sheetLength; i++) {
    sheetRow = string.slice(i, i + 1);
    dataRow = string.slice(dataRange[1] + times - 1, dataRange[1] + times);
    query = '=IMPORTRANGE("' + key + '", ' + sheetRow +'2&"!' + dataRow + (dataRange[0] + 5) + ':' + dataRow + (dataRange[0] + dataRange[2] + 5) + '")';
    result.push(query);
  }
  
  return result;
}

function setDataQuery(currentSheet, queryArr, reportMetrics, times) {
  var metrics = reportMetrics[1].split(/\r\n|\r|\n/);

  var setQueryArr = [];
  setQueryArr.push(queryArr);

  currentSheet.getSheetByName(metrics[times]).getRange(3, 1, 1, queryArr.length).setValues(setQueryArr);
}


/*
// 旧バージョンなので一旦コメントアウト
function setQuery(currentSheet, queryArr, reportMetrics, times) {
  var metrics = reportMetrics[1].split(/\r\n|\r|\n/);

  var setQueryArr = [];
  setQueryArr.push(queryArr);

  currentSheet.getSheetByName(metrics[times]).getRange(2, 1, 1, queryArr.length).setValues(setQueryArr);
}

// 旧バージョンなので一旦コメントアウト
function getQueryArr(key, dataRange, sheetLength, times) {
  var result = [];
  var string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var sheetRow = '';
  var dataRow = '';
  var query = '';

  // sheetLength = 25;

  // =IMPORTRANGE("1_xTo2_DikUJJuJdVmZGEwEmA9YavJih_SXtOJi-B2dQ",A1 & "!C16:C27") を作成し、resultへpushする
  for (var i = 0; i < sheetLength; i++) {
    sheetRow = string.slice(i, i + 1);
    dataRow = string.slice(dataRange[1] + times - 1, dataRange[1] + times);
    query = '=IMPORTRANGE("' + key + '", ' + sheetRow +'1&"!' + dataRow + dataRange[0] + ':' + dataRow + (dataRange[0] + dataRange[2]) + '")';
    result.push(query);
  }
  
  return result;
}
*/

// GAによって出力されてdataRangeを取得する。
// 配列→スプレッドシート　への置換のためにスクリプトがちょっと汚い。
// マジックナンバーで+1する箇所が多くて困惑。
function getDataRange(ss, sheetName) {
  var spreadSheetData = ss.getSheetByName(sheetName).getDataRange().getValues();
  var dataStartRow = 10;
  var topRow = 11, topColumn = 0, lastRow = 0, lastColumn = 0;

  lastRow = spreadSheetData.length - dataStartRow;

  var i = 0;
  do {
    i++;
    if (spreadSheetData[dataStartRow][i] != "") {
      break;
    }
  } while(spreadSheetData[dataStartRow].length > i);

  // 配列からシートへ置き換える
  topColumn = i + 1;
  lastColumn = spreadSheetData[dataStartRow].length - topColumn + 1;

  return [topRow, topColumn, lastRow, lastColumn];
}


// データのヘッダーを入力する
function setHeader(currentSheet, reportNames, reportMetrics) {
  var metrics = reportMetrics[1].split(/\r\n|\r|\n/);

  // 入力する配列 reportNamesを入力できる形に成形
  reportNames.shift();
  var setArr = [];
  setArr.push(reportNames);
  var range;

  for (var i = 0; i < metrics.length; i++) {
    // 空行チェック
    if(metrics[i] == "") continue;
    
    range = currentSheet.getSheetByName(metrics[i]).getRange(2, 1, 1, reportNames.length);
    range.setValues(setArr);  
    range.setBackground("#d9d9d9");    
  }
}


// 入力配列に沿ったシートを作成する
// 最初のデータのMetricsに沿ったシート作成をします！！
function makeSpreadSheet(reportMetrics, currentSheet) {

  var existSheetNames = wmap_getSheetsName();
  var metrics = reportMetrics[1].split(/\r\n|\r|\n/);

  var isMatch;
  for(var i = 0; i < metrics.length; i++) {
    // 空行チェック
    if(metrics[i] == "") continue;
    
    isMatch = false;
    for (var j = 0; j < existSheetNames.length; j++) {
      if(metrics[i] == existSheetNames[j]) {
        currentSheet.getSheetByName(metrics[i]).clear();
        isMatch = true;
        break;
      }
    }
    if(!isMatch) {
      currentSheet.insertSheet(metrics[i]);
    }
  }
}


// 既存のシート名一覧を取得する
function wmap_getSheetsName(){
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      sheet_names.push(sheets[i].getName());
    }
  }
  return sheet_names;
}


// 未完成
// 27→AAに置換
// 一旦放置　やっぱうまいごと思いつかないorz　人に頼ってもいいやもう。
function replaceNumToString() {
  var num = 677;
  var string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var variety = 26; //A〜Zの文字列の種類
  var result = '';
  
  var i = 0;
  do {
    i++
  } while(num >= Math.pow(variety, i));

  for(; i > 0; i--) {
    for (var j = 0; j < variety; j++) {
//      Logger.log("[i, j] = [" + i + ", " + j + "]");
      Logger.log((Math.pow(variety, i - 1) * j) + " <= " + num + " < " + (Math.pow(variety, i - 1) * (j + 1)));
      if ((Math.pow(variety, i - 1) * j) <= num && num < (Math.pow(variety, i - 1) * (j + 1))) {
        result = result + string.substr(j - 1, 1);
        num = num - Math.pow(variety, (i - 1)) * j;
//        Logger.log("[i, j] = [" + i + ", " + j + "]");
        break;      
      } 
    }
  }

  Logger.log(num);
  Logger.log(result);

}