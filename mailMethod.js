function mailData(from, date, subject, attachmentName, attachmentDataString) {
  this.from = from;
  this.date = date;
  this.subject = subject;
  this.attachmentName = attachmentName;
  this.attachmentDataString = attachmentDataString;
}

//對同個寄件人的對話群組進行解析
function getMailData(messages) {
  let lastMessage = messages[messages.length - 1];

  //寄件人、信件標題、寄件時間、附加檔案
  let iFrom = lastMessage.getFrom();
  let iSubject = lastMessage.getSubject();
  let iDate = Utilities.formatDate(lastMessage.getDate(), 'GMT+8', 'yyyy-MM-dd_HH:mm');
  let iAttachments = lastMessage.getAttachments();

  let mailDataArray = [];

  //如果信件內文有speedtest資料，就擷取資料，並放到要輸出的資料隊列去
  let iBody = lastMessage.getBody();
  let iBodyIndex = iBody.indexOf('￦ﾗﾥ￦ﾜﾟ');
  if(iBodyIndex>=0){
    let iBodyTable = iBody.substring(iBodyIndex).replaceAll("&quot;","\"").replaceAll(/<[^>]*>/g,"");
    let iMailData = new mailData(iFrom, iDate, iSubject, "SpeedtestPutMailBodyData", iBodyTable);
    mailDataArray.push(iMailData);
  }

  //如果附加檔案檔名符合特徵，就擷取資料，並放到要輸出的資料隊列去
  for (let i in iAttachments) {
    let jName = iAttachments[i].getName();
    let jDataAsString = iAttachments[i].getDataAsString();

    let iMailData = new mailData(iFrom, iDate, iSubject, jName, jDataAsString);
    mailDataArray.push(iMailData);
  }

  return mailDataArray;
}