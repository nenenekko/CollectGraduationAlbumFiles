const large_classes = ["研究室","サークル"];
const classes = [["Ⅰ類","Ⅱ類","Ⅲ類"],["同好会系","文化系","体育会系"]];
const sub_names = ["指導教員","サークル"]


function myFunction() {
  // DriveApp.getRootFolder()
}


function Init(folder_id){
  var sp_sheet = SpreadsheetApp.create("提出ファイルの管理");
  var sp_file = DriveApp.getFileById(sp_sheet.getId());
  FileMigration(folder_id,sp_file);
  CreateGoogleForms(sp_sheet,folder_id);
  AddSheet(sp_sheet);
  CreateManageStatusSheets(folder_id);
  return sp_sheet.getId();
}


function CreateGoogleForms(sp_sheet,folder_id){
  var forms = [];
  forms.push(CreateGoogleForm(large_classes[0],classes[0],folder_id));
  forms.push(CreateGoogleForm(large_classes[1],classes[1],folder_id));
  var binded_sheets = BindForms(forms,sp_sheet,large_classes);
  for(var i=0; i < binded_sheets.length; i++){
    binded_sheets[i].getRange(1,10,1,2).setValues([['最新版でなければ✖','目黒会側の確認の有無']]);
    
    const values = ['未確認','ダウンロード及び保存済み','提出の確認はしました'];
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
    for(var j=1; j<=450; j++)
      binded_sheets[i].getRange(2,11,j,1).setDataValidation(rule);
  }
}


function CreateGoogleForm(form_name,class_list,folder_id){
  var [student_name,team_name] = GetNames(form_name,class_list);
  form = FormApp.create("ファイル提出(" + form_name + ")"); //フォーム作成
  form.addListItem().setTitle("担当アルバム委員の名前").setChoiceValues(student_name).setRequired(true); //プルダウンリスト
  jump_item = form.addListItem().setTitle('担当する団体の分類').setRequired(true); //ラジオボタン
  pages = []
  for(var i = 0; i < class_list.length ; i++){
    pages.push(form.addPageBreakItem().setTitle(class_list[i] + '用'));
    form.addListItem().setTitle(class_list[i] + ':' + form_name + "名").setChoiceValues(team_name[i]).setRequired(true); //プルダウンリスト
  }
  pages.push(form.addPageBreakItem().setTitle('Giga fileのリンク提出'));
  form.addTextItem().setTitle('Giga fileのリンク').setRequired(true); //記述式
  form.addTextItem().setTitle('削除キー').setRequired(true);
  form.addMultipleChoiceItem().setTitle('サブプランであるか').setChoiceValues(['サブプランです']).setHelpText('提出する写真はサブプランですか？(GimpやPhotshopで編集されたものではないならサブプラン)\n*サブプランではない場合，回答しない');
  candidates = []
  for(var i = 1; i <= class_list.length ; i++){
    pages[i].setGoToPage(pages[class_list.length]);
    candidates.push(jump_item.createChoice(class_list[i-1], pages[i-1]));
  }
  jump_item.setChoices(candidates);
  const formFile = DriveApp.getFileById(form.getId());
  FileMigration(folder_id,formFile)
  return form
}


function GetNames(form_name,team_class){
  utilitySheet = SpreadsheetApp.openById(utlity_sheet_id);
  student_sheet = utilitySheet.getSheetByName("アルバム委員");
  student_name = Sheet2Name(student_sheet);
  team_name = []
  if(form_name == "研究室"){
    for(var i=0; i<team_class.length; i++)
      team_name.push(Sheet2Name(utilitySheet.getSheetByName(team_class[i] + sub_names[0])));
  }else if(form_name == "サークル"){
    for(var i=0; i<team_class.length; i++)
      team_name.push(Sheet2Name(utilitySheet.getSheetByName(team_class[i] + sub_names[1])));
  }
  return [student_name,team_name]
}


function Sheet2Name(sheet){
  name_array = []
  cell = sheet.getDataRange().getValues();
  for(var i = 1; i < cell.length; i++)
    if(cell[i] != "")
      name_array.push(cell[i])
  return name_array
}


function AddSheet(sp_sheet){
  CreateTableSheet(sp_sheet,large_classes[0],classes[0],sub_names[0]);
  CreateTableSheet(sp_sheet,large_classes[1],classes[1],sub_names[1]);
  var setting_sheet = sp_sheet.insertSheet("setting");
  setting_sheet.getRange(1,1,2,3).setValues([['現在の最新セル(研究室)','','現在の最新セル(サークル)'],[0,'',0]]);
  var sheets = sp_sheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('シート') != -1) {
        sp_sheet.deleteSheet(sheets[i]);
        break;
    }
  }
  return sp_sheet;
}


function DeleteSheet(sp_sheet,name){
  var sheets = sp_sheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf(name) != -1) {
        sp_sheet.deleteSheet(sheets[i]);
        break;
    }
  }
}


function CreateTableSheet(sp_sheet,name,index_name,sub_name){
  var sheet_name = name + "提出状況一覧";
  var table_sheet = sp_sheet.insertSheet(sheet_name);
  table_sheet.getRange(1,1,1,6).setValues([['',name + "名",'担当アルバム委員','GigaFileリンク','削除キー','掲載辞退']]);
  utilitySheet = SpreadsheetApp.openById(utlity_sheet_id);
  last_team_size = 0
  for(var i=0; i < index_name.length; i++){
    team_sheet = utilitySheet.getSheetByName(index_name[i] + sub_name);
    team_name = Sheet2Name(team_sheet);
    for(var j=0; j < team_name.length; j++){
      if(j != 0)
        table_sheet.getRange(j+2+last_team_size,1,1,2).setValues([["",team_name[j]]]);
      else
        table_sheet.getRange(j+2+last_team_size,1,1,2).setValues([[index_name[i],team_name[j]]]);
    }
    last_team_size += team_name.length;
  }
  
  //プルダウンの選択肢を配列で指定
  const values = ['掲載辞退','期限延長'];
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  //リストをセットするセル範囲を取得
  for(var i=1; i<=last_team_size; i++)
    table_sheet.getRange(2,6,i,1).setDataValidation(rule);
}


function BindForms(forms,sp_sheet,name){
  var sheet_name = [];
  for(var i = 0; i < name.length; i++)
    sheet_name.push("ファイル提出("+ name[i] + ")");
  for(var i = 0; i < forms.length; i++)
    forms[i].setDestination(FormApp.DestinationType.SPREADSHEET, sp_sheet.getId());
  var binded_sheets = []
  var sheets = sp_sheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('Form') != -1 || name.indexOf('フォーム') != -1){
      if(name.indexOf('1') != -1)
        sheets[i].setName(sheet_name[0]);
      else if(name.indexOf('2') != -1)
        sheets[i].setName(sheet_name[1]);
      binded_sheets.push(sheets[i]);
    }
  }
  return binded_sheets;
}


function FileMigration(folder_id,file){
  DriveApp.getFolderById(folder_id).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}


function CreateManageStatusSheets(folder_id){
  var sp_sheet = SpreadsheetApp.create("状況確認シート");
  for(var i = 0; i < large_classes.length; i++){
    for(var j = 0; j < classes[i].length; j++){
      CreateManageStatusSheet(sp_sheet,large_classes[i],classes[i][j],sub_names[i]);
    }
  }
  var sheets = sp_sheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('シート') != -1) {
      sp_sheet.deleteSheet(sheets[i]);
      break;
    }
  }
  FileMigration(folder_id,DriveApp.getFileById(sp_sheet.getId()));
}


function CreateManageStatusSheet(sp_sheet,large_class_name,class_name,sub_name){
  var table_sheet = sp_sheet.insertSheet(class_name + "状況確認シート(" + large_class_name + ")");
  table_sheet.getRange(1,1,1,10).setValues([[large_class_name + "名",'担当アルバム委員','代表者名','Slack加入','①団体に連絡','②代表者のメアド確保&詳細連絡','③〆切確認メール送信','④写真回収','⑤写真提出','備考']]);
  utilitySheet = SpreadsheetApp.openById(utlity_sheet_id);
  class_sheet = utilitySheet.getSheetByName(class_name + sub_name);
  class_member_name = Sheet2Name(class_sheet);
  committee_sheet = utilitySheet.getSheetByName("アルバム委員");
  committee_member_name = Sheet2Name(committee_sheet);
  for(var i=0; i < class_member_name.length; i++)
    table_sheet.getRange(i+2,1,1,1).setValue(class_member_name[i]);
  const status_values = ['未','済','やってる途中','難航中','相談したい'];
  const status_rule = SpreadsheetApp.newDataValidation().requireValueInList(status_values).build();
  const slcak_values = ['有','無'];
  const slack_rule = SpreadsheetApp.newDataValidation().requireValueInList(slcak_values).build();
  const committee_rule = SpreadsheetApp.newDataValidation().requireValueInList(committee_member_name).build();
  //リストをセットするセル範囲を取得
  for(var i=1; i<=class_member_name.length; i++){
    table_sheet.getRange(2,5,i,5).setDataValidation(status_rule);
    table_sheet.getRange(2,4,i,1).setDataValidation(slack_rule);
    table_sheet.getRange(2,2,i,1).setDataValidation(committee_rule);
  }
}