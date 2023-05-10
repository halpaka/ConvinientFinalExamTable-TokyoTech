function myFunction(e) {
  //上限に達したらフォームを閉じる
  if(MailApp.getRemainingDailyQuota() == 0){
    const form = FormApp.getActiveForm();
    form.setAcceptingResponses(false);
    return;
  }
  
  //下の値は更新時の最新期末予定表
  const year = 2023;  //年
  var month = 2;  //月
  const firstday = 3; //初日

  var i,j;

  //期末試験情報が書かれたスプレッドシートを取得
  const sheet = SpreadsheetApp.openByUrl('スプレッドシートのURL').getSheetByName('シート名'); //スプレッドシート新規作成に合わせ変更
  const sheetcolumn = sheet.getLastColumn();
  const sheetrow = sheet.getLastRow();
  const sheetdata = sheet.getRange(1,1,sheetrow,sheetcolumn-2).getValues(); //スプレッドシート内のデータを取得

  //メアド取得
  const email = e.response.getRespondentEmail();
  const titmail = email.indexOf('@m.titech.ac.jp');
  if(titmail < 0)　return;  //大学メアド以外ははじく

  //フォームの内容取得
  const anscode = e.response.getItemResponses();
  const codemax = Number(anscode[0].getResponse());
  let sbjcode = [];
  for(i = 0; i < codemax; i++) sbjcode[i] = anscode[i+1].getResponse();

  //送信するデータをまとめる変数
  var sendcode = [];
  var sendname = [];
  var senddate = [];
  var sendtest = [];
  var sendstart = [];
  var sendfinish = [];

  //科目検索
  var time = [];
  for(i = 0; i < codemax; i++){
    for(j = sheetrow-1; j >= 0; j--){
      if(sheetdata[j][3] == sbjcode[i]){
        sendcode[i] = sheetdata[j][3];
        sendname[i] = sheetdata[j][4].substr(0,sheetdata[j][4].indexOf('\n'));
        senddate[i] = sheetdata[j][0];
        if(sheetdata[j][1].indexOf('期末試験') != -1) sendtest[i] = '【期末試験】';
        else sendtest[i] = '【最終授業】';
        time = time_period(sheetdata[j][2]);
        sendstart[i] = time[0];
        sendfinish[i] = time[1];
        break;
      }
    }
    if(j == -1){
      sendcode[i] = sbjcode[i];
      sendtest[i] = '0';
    }
  }

  //csv用データ編集
  var csvsbj = [];
  var csvday = [];
  var csvstarttime = [];
  var csvendtime = [];
  var daybox;
  for(i = 0; i < codemax; i++){
    if(sendtest[i] != '0'){
      csvsbj[i] = sendname[i] + sendtest[i];
      csvday[i] = String(year) + '/';
      daybox = senddate[i].substr(0,senddate[i].indexOf('('));
      if(firstday > Number(daybox)) csvday[i] += String(month+1) + '/' + daybox;
      else csvday[i] += String(month) + '/' + daybox;
      csvstarttime[i] = sendstart[i];
      csvendtime[i] = sendfinish[i];
    }
    else csvsbj[i] = '0';
  }

  //csvファイル作成
  var csv = '"Subject","Start Date","Start Time","End Date","End Time"\r\n';
  for(i = 0; i < codemax; i++){
    if(csvsbj[i] != '0'){
      csv += '"' + csvsbj[i] + '",';
      csv += '"' + csvday[i] + '",';
      csv += '"' + csvstarttime[i] + '",';
      csv += '"' + csvday[i] + '",';
      csv += '"' + csvendtime[i] + '"\r\n';
    }
  }
  const csvName = 'TermEndExam-' + String(year) + '-' + String(month) + '-' + String(firstday) + '.csv';
  var blob = Utilities.newBlob("", 'text/comma-separated-values', csvName);
  blob.setDataFromString(csv, "UTF-8");

  //メールの本文作成
  let text = '科目コードによる試験日の検索結果です。必ず実際のものと合致しているか改めてご確認ください。<br/><br/><br/>';
  for(i = 0; i < codemax; i++){
    text += '科目コード : ' + sendcode[i] + '<br/>';
    if(sendtest[i] != '0'){
      text += '科目名 : ' + sendname[i] + '<br/>';
      text += '形態 : ' + sendtest[i] + '<br/>';
      text += '日付 : ' + csvday[i] + '<br/>';
      text += '時間 : ' + sendstart[i] + ' ~ ' + sendfinish[i] + '<br/>';
      text += '<br/><br/>'
    }
    else{
      text += 'お探しの科目コードは見つかりませんでした。<br/>';
      text += '期末試験表のPDFに記載のない授業か科目コードに誤りがあるかのどちらかだと思われます。<br/>';
      text += '<br/><br/>';
    }
  }
  text += 'csvファイルに対応しているカレンダーアプリ（Googleカレンダー）にインポートするとカレンダーに反映されます。';
  text += '結果が正しく反映されているかご確認ください。';

  var options = {htmlBody: text,attachments:[blob]};
  MailApp.sendEmail(email, '期末試験日程検索結果', text, options);
}

//開始終了時間の取得
function time_period(period){
  const start_time = ['8:50','9:40','10:45','11:35','13:45','14:35','15:40','16:30','17:30'];
  const end_time = ['10:30','11:35','12:25','14:35','15:25','16:30','17:20','18:20','19:10'];
  const period_start_list = ['1-','2-','3-','4-','5-','6-','7-','8-','9-'];
  const period_end_list = ['-2','-3','-4','-5','-6','-7','-8','-9','-10'];
  var start,end;
  for(var i = 0; i < 9; i++){
    if(period.indexOf(period_start_list[i]) != -1) start = start_time[i];
    if(period.indexOf(period_end_list[i]) != -1) end = end_time[i];
  }
  return [start,end];
}
