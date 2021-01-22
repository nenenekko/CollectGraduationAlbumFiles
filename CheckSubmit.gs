//スプレッドシートを取得
var utilitySheet;
var spSheet;

/*各シートを取得*/
var lab_sheet;
var circle_sheet;
var setting_sheet;
var lab_sort_sheet;
var circle_sort_sheet;
var mail_sheet;

lab_form_count = 0
circle_form_count = 0

/*メアド管理*/
var meguro_mails;
var committee_mails;
var mails;

function GetSettings(){
  //スプレッドシートを取得
  var project_info_sheet = SpreadsheetApp.openById(project_info_sheet_id);
  var pl_sheet = project_info_sheet.getSheetByName("プロジェクトログ");
  var log_cell = pl_sheet.getDataRange().getValues();
  var my_sp_sheet_id = log_cell[log_cell.length - 1][2]
  spSheet = SpreadsheetApp.openById(my_sp_sheet_id);
  utilitySheet = SpreadsheetApp.openById(utlity_sheet_id);

  /*各シートを取得*/
  lab_sheet = spSheet.getSheetByName("ファイル提出(研究室)");
  circle_sheet = spSheet.getSheetByName("ファイル提出(サークル)");
  setting_sheet = spSheet.getSheetByName("setting");
  lab_sort_sheet = spSheet.getSheetByName("研究室提出状況一覧");
  circle_sort_sheet = spSheet.getSheetByName("サークル提出状況一覧");

  mail_sheet = utilitySheet.getSheetByName("担当メアド");
  mail_cell = mail_sheet.getDataRange().getValues();
  meguro_mails = [[],[]];
  committee_mails = [[],[]];
  for(var i = 2; i < mail_cell.length; i++){
    if(mail_cell[i][0] != ""){
      meguro_mails[0].push(mail_cell[i][0])
      meguro_mails[1].push(mail_cell[i][1])
    }
    if(mail_cell[i][2] != ""){
      committee_mails[0].push(mail_cell[i][2])
      committee_mails[1].push(mail_cell[i][3])
    }
  }
  mails = [meguro_mails,committee_mails]
  
  settings = setting_sheet.getDataRange().getValues();
  Logger.log(settings);
  lab_form_count = settings[1][0];
  circle_form_count = settings[1][2];
}

/*一定時間ごとにこいつを動かす*/
function CheckNewFileAndSendMail(){
  GetSettings()
  new_lab_file_count = CheckNewFile(lab_form_count,lab_sheet,lab_sort_sheet);
  new_circle_file_count = CheckNewFile(circle_form_count,circle_sheet,circle_sort_sheet);
  
  lab_form_count += new_lab_file_count;
  circle_form_count += new_circle_file_count;
  
  if(new_lab_file_count != 0 || new_circle_file_count != 0){
    Send(new_lab_file_count,new_circle_file_count,mails);
    setting_sheet.getRange(2,1,1,1).setValue(lab_form_count);
    setting_sheet.getRange(2,3,1,1).setValue(circle_form_count);
  }
}

function CheckNewFile(count,submited_sheet,sort_sheet){
  new_file_count = 0
  const data = submited_sheet.getDataRange().getValues();
  for(var i = count+1; i < data.length; i++){
    if(data[i][0] == "")
      break;
    CheckReSubmit(data,i,submited_sheet)
    EditTotalTable(data[i],sort_sheet)
    new_file_count++;
  }
  return new_file_count;
}

function GetTarget(new_data){
  var target_name = "";
  var index = -1;
  if(new_data[3] != ""){
    target_name = new_data[3];
    index = 3;
  }else if(new_data[4] != ""){
    target_name = new_data[4];
    index = 4;
  }else if(new_data[5] != ""){
    target_name = new_data[5];
    index = 5;
  }
  target_name.replace('　',' ')
  return [target_name,index];
}

function CheckReSubmit(data,new_data_index,submited_sheet){
  var [target_name, name_index] = GetTarget(data[new_data_index]);
  for(var i = new_data_index - 1; i >= 0 ; i--){
    if(data[i][name_index].replace('　',' ') == target_name){
      submited_sheet.getRange(i+1,10,1,1).setValue("✖");
      break;
    }
  }
}

function EditTotalTable(new_data,sort_sheet){
  var [target_name, _] = GetTarget(new_data);
  var sort_data = sort_sheet.getDataRange().getValues();
  for(var i=0; i<sort_data.length;i++){
    if(sort_data[i][1].replace('　',' ') == target_name){
      var renew_data = []
      renew_data.push(new_data[1])
      renew_data.push(new_data[6])
      renew_data.push(new_data[7])
      sort_sheet.getRange(i+1,3,1,1).setValue(renew_data[0]);
      sort_sheet.getRange(i+1,4,1,1).setValue(renew_data[1]);
      sort_sheet.getRange(i+1,5,1,1).setValue(renew_data[2]);
    }
  }
}

function Send(new_lab_count,new_circle_count,mails) {
  const recipient = mails[0][1][0]; //送信先のメールアドレス 
  const subject = '新規提出ファイルがあります';

  const recipientCompany = '目黒会';
  var name_str = mails[0][0][0] + 'さん,';
  var cc_str = "";
  for(var i = 1; i < mails[0][0].length ; i++){
    cc_str += mails[0][1][i] + ',';
    name_str += mails[0][0][i] + 'さん';
    if(i != mails[0][0].length -1){
      cc_str += ',';
      name_str += ',';
    }
  }
  for(var i = 0; i < mails[1][1].length ; i++){
    cc_str += mails[0][1][i]
    if(i != mails[1][1].length -1){
      cc_str += ','
    }
  }
  const body = `${recipientCompany}  ${name_str}\n\n研究室関連の新規ファイルが${new_lab_count}件,サークル関連の新規ファイルが${new_circle_count}件届いています．ご確認をお願い致します.\n\n*本メールはbotでの送信です．ご不明点等ございましたら山根(nekko0429@gmail.com)までお願い致します．`;
  const options = {name: '山根大輝(自動送信bot)', cc: cc_str};
  
  GmailApp.sendEmail(recipient, subject, body, options);
}