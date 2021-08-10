//openSignal & speedTest 的標題對應表
const comparisonTableList = {
  'dl_speed': 'DL',
  'ul_speed': 'UL',
  'my_lat': 'Latitude',
  'my_lon': 'Longitude',
  'CID': 'EnbID&CellID',
  'timestamp': ['Date', 'Time'],
  '日期': ['Date', 'Time'],
  '緯度': 'Latitude',
  '經度': 'Longitude',
  '下載速度': 'DL',
  '上傳速度': 'UL',
  '下載': 'DL',
  '上載': 'UL',
  '￦ﾗﾥ￦ﾜﾟ': ['Date', 'Time'],
  '￧ﾷﾯ￥ﾺﾦ': 'Latitude',
  '￧ﾶﾓ￥ﾺﾦ': 'Longitude',
  '￤ﾸﾋ￨ﾼﾉ': 'DL',
  '￤ﾸﾊ￥ﾂﾳ': 'UL'
}


//輸出到 sheet 的標題
const outputTableTitle = ['Date', 'Time', 'DL', 'UL', 'Latitude',
  'Longitude', 'EnbID&CellID', 'mailFrom', 'mailAttFileName', 'mailSubject', 'mailDate'];


//建立sheet物件並查詢，如果找不到物件，就建立一個新的，放在最尾端
function createWorkSheet(workSheetName) {
  let sheetObject = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetObject.getSheetByName(workSheetName) === null) {
    sheetObject.insertSheet(workSheetName, sheetObject.getSheets().length);
  }
}


//寫入myTable資料，到指定工作表
function writeWorksheet(workSheetName, writeDatas) {
  let worksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(workSheetName);

  //表內沒資料，添加標題後再增加資料
  if (worksheet.getRange("A1").isBlank()) {
    worksheet.appendRow(outputTableTitle);
  }

  //表內有資料，直接安插資料在最後一排
  for (const writeData of writeDatas) {
    worksheet.appendRow(writeData);
  }
}


let modifyMailData = function (key, value) {

  let enbId_CellId = function (myValue) {
    const hex = 256;
    let enbID = (Math.floor(myValue / hex)).toString();
    let cellID = (myValue % hex).toString();
    return enbID + "/" + cellID;
  }

  let UTCDate = function (myValue) {
    let ret = new Date(parseInt(myValue));
    let retArray = {};
    retArray['Date'] = Utilities.formatDate(ret, 'GMT+8', 'yyyy/MM/dd');
    retArray['Time'] = Utilities.formatDate(ret, 'GMT+8', 'HH:mm');
    return retArray;
  }

  let strDate = function (myValue) {
    let ret = myValue.split(" ");
    let retArray = {};
    retArray['Date'] = ret[0].replace(/-/gi, '/');
    retArray['Time'] = ret[1];
    return retArray;
  }

  let bitToMbps_check = function (myValue) {
    if (myValue.toString().indexOf('.') > 0) {
      return myValue;
    }
    return myValue / 1000 / 1000 * 8;
  }

  let bpsToMbps = function (myValue) {
    return myValue / 1000;
  }

  let remainValue = function (myValue) {
    return myValue;
  }

  let removeValueQuotStrDate = function (myValue) {
    let ret = myValue.replaceAll("\"","").split(" ");
    let retArray = {};
    retArray['Date'] = ret[0].replace(/-/gi, '/');
    retArray['Time'] = ret[1];
    return retArray;
  }

  let removeValueQuot = function (myValue) {
    return myValue.replaceAll("\"","")
  }
  
  //Object-literal 對應區域
  let actions = {
    //opensignal區域
    'dl_speed': bpsToMbps,
    'ul_speed': bpsToMbps,
    'my_lat': remainValue,
    'my_lon': remainValue,
    'CID': enbId_CellId,
    'timestamp': UTCDate,
    //speedtest區域
    '日期': strDate,
    '下載速度': bitToMbps_check,
    '上傳速度': bitToMbps_check,
    '緯度': remainValue,
    '經度': remainValue,
    //較早版本的speedtest用詞
    '下載': bpsToMbps,
    '上載': bpsToMbps,
    //更早版本(2016before)的speedtest亂碼用詞
    '￦ﾗﾥ￦ﾜﾟ': removeValueQuotStrDate,
    //緯度
    '￧ﾷﾯ￥ﾺﾦ': removeValueQuot,
    //經度
    '￧ﾶﾓ￥ﾺﾦ': removeValueQuot,
    //下載
    '￤ﾸﾋ￨ﾼﾉ': removeValueQuot,
    //上傳
    '￤ﾸﾊ￥ﾂﾳ': removeValueQuot
  }
  //假如沒有同名對應方法便throw Error，否則呼叫對應actions
  if (typeof actions[key] !== 'function') {
    throw new Error("Action not found.");
  }
  return actions[key](value);
}


//製作自訂選單到試算表內，並連結函式(請勿修改函示名稱，GAS指定的接口)
function onOpen() {
  let menu = [{
    name: "讀取mail資料",
    functionName: "main"
  }];

  SpreadsheetApp.getActiveSpreadsheet().addMenu("擴充功能選單", menu);
}