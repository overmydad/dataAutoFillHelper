//將map物件，轉換為array ["key:value"]，方便增加查詢條件 
let iMap = {
  toString: function (_obj) {
    let iArray = [];
    for (let i in _obj) {
      if (_obj.hasOwnProperty(i)) {
        iArray.push(i + ":" + _obj[i]);
      }
    }
    return iArray.join(" ");
  }
}

let mailQuery = {
  'subject': 'Speedtest 結果（CSV 匯出）'
}

function main() {
  //取得查詢結果的陣列
  let mails = GmailApp.search(iMap.toString(mailQuery));

  //製造一個新工作表，並保留工作表名稱
  let worksheetName = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd_HH:mm');
  createWorkSheet(worksheetName);

  //對每一個寄件人群組的郵件進行資料提取，整理後，寫入到試算表。
  for (const mail of mails) {
    let iData = getMailData(mail.getMessages());
    writeWorksheet(worksheetName, transformDataFormat(iData));
  }
}