// 作成するシート名一覧が入っているドキュメントのキー
var docKey = "1jYo0A8GSaraot0iQYTI7F2DvKfeP4LY1Dl3o_Jqbm2s";
// シートを作りたいディレクトリキー
var folderKey = "1Wk921VWYQihEfPHkXmFSQItiiOqESPBF";


function makeSpreadSheet() {
  var ssNames = getNewSSNames();
  
  var IDs = makeSheets(ssNames);

  Logger.log(IDs);
}


function makeSheets(ssNames) {
  var folder = DriveApp.getFolderById(folderKey);
  var IDs = [];  
  var ssId, file;
  
  for (var i = 0; i < ssNames.length; i++) {  
    if (folder.getFilesByName(ssNames[i]).hasNext()) {
      Logger.log(ssNames[i]+"があります");
      ssId = DriveApp.getFilesByName(ssNames[i]).next().getId();
      IDs.push(ssId);
    } else {
      Logger.log(ssNames[i]+"を作成");
      ssId = SpreadsheetApp.create(ssNames[i]).getId();
      file = DriveApp.getFileById(ssId);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      IDs.push(ssId);
    }
  }
  
  return IDs;  
}



function makeNewSS(ssNames) {
  var idSS = [];
  var id;

  for (var i = 0; i < ssNames.length; i++) {
    id = SpreadsheetApp.create(ssNames[i]).getId();
    idSS.push(id);
  }  
  
  Logger.log(idSS);
}


function getNewSSNames() {
  var docTest = DocumentApp.openById(docKey);
  var textBody = docTest.getBody().getText();
  
  var metrics = textBody.split(/\r\n|\r|\n/);
  
  return metrics;
}
