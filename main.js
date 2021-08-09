/*
=======================
主運行函式
google apps script limit
https://docs.google.com/macros/dashboard
gmail相關參數
https://developers.google.com/apps-script/reference/gmail/
=======================
*/
function mainFunction() {
  //將myConfig.gs的gmail查詢式物件轉換為字串
  var tempArray = myObj.toString(gmailQuery);
  //查詢結果的信件集合
  var mails = GmailApp.search(tempArray.join(" "));
  //製造一個新工作表，並保留工作表名稱
  var worksheetName = makeTabs();
  for (var i = 0; i < mails.length; i++) {
    var messages = mails[i].getMessages();
    //有信件重複寄的話，直接取最後(新)一封(不管是寄重複，回覆信件...等等都算)
    var j = messages.length - 1;
    //寄件人，信件標題，寄件時間
    var from = messages[j].getFrom();
    var subject = messages[j].getSubject();
    var mailDate = messages[j].getDate();
    //信件的附加檔案
    var attachments = messages[j].getAttachments();
    var csvData = [];
    var csvTitle = [];
    var csvMailINFO = [];
    for (var k = 0; k < attachments.length; k++) {
      //附加檔案的filename
      var fileName = attachments[k].getName();
      //所有檔案儲存成陣列
      var fileContent = attachments[k].getDataAsString();

      //用附加檔案的filename，判斷要添加什麼資料到各類陣列裡
      if (fileName.indexOf('results.csv') === 0) {
        csvData.push(fileContent);
        csvTitle.push(speedTestTitle);
        csvMailINFO.push({
          'mailForm': from,
          'mailAttFilename': fileName,
          'mailSubject': subject,
          'mailDate': mailDate
        });
      } else if (fileName.indexOf('osm_speedtests_') === 0) {
        csvData.push(fileContent);
        csvTitle.push(openSignalTitle);
        csvMailINFO.push({
          'mailForm': from,
          'mailAttFilename': fileName,
          'mailSubject': subject,
          'mailDate': mailDate
        });
      }
    }
    //將資料陣列與其他資訊，輸出為一張完整的陣列表單，並寫入到資料表裡面
    writeTabs(worksheetName, myCsv(csvData, csvTitle, csvMailINFO));
  }
}