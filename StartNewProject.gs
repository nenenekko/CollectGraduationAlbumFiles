function CreateNewProject(){
  var project_log_sheet = SpreadsheetApp.openById(project_info_sheet_id);
  var pl_sheet = project_log_sheet.getSheetByName("プロジェクトログ");
  var log_cell = pl_sheet.getDataRange().getValues();
  var folder_id = CreateFolder(log_cell[log_cell.length - 1][1]);
  var created_sheet_id = Init(folder_id);
  pl_sheet.getRange(log_cell.length,3,1,1).setValue(created_sheet_id);
}

function CreateFolder(folder_name){
  var folder_id = root_folder_id;　//フォルダを作成する場所（フォルダ）のGoogleドライブ上のIDを指定
  var folder = DriveApp.getFolderById(folder_id);　//IDからフォルダを取得
  var new_folder = folder.createFolder(folder_name);　//変数folderで取得した場所に、指定名称でフォルダ作成
  var new_folder_id = new_folder.getId();　//先ほど作成したフォルダのIDを取得
  return new_folder_id;
}