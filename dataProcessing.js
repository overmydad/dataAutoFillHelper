let transformOutputData = [];

//將原始csv檔案提取需要的資料，轉換到另外一個二維陣列上
function transformDataFormat(datas) {
  transformOutputData = [];
  //附件檔案層級
  for (const data of datas) {
    updateCsvData(data, transformOutputData);
  }
  return transformOutputData;
}

function updateCsvData(mailData, output) {

  let iFrom = mailData.from;
  let iDate = mailData.date;
  let iSubject = mailData.subject;
  let iName = mailData.attachmentName;
  let iTitle = selectSheetTilte(iName);
  //Utilities.parseCsv output is String 2D array
  let iDatas = Utilities.parseCsv(mailData.attachmentDataString);

  let iTitleIndex = {};

  //附件檔案裡面的列數
  for (const iData of iDatas) {
    let tempData = [];

    //遇到標題欄位時，儲存要擷取的標題，跟附加檔案裡標題出現的位置index。
    if (iData.indexOf(iTitle[0]) != -1) {
      iTitleIndex = setTitleIndex(iTitle, iData);
      continue;
    }

    //指定的第一個欄位如果為空格，即停止輸出(數值處理沒做檢查)
    if (iData[iTitleIndex[iTitle[0]]] === ""){
      continue;
    }
    
    //將iTitleIndex取得的index反向搜尋數據，並按照指定位置儲存
    for (const ikey in iTitleIndex) {
      let iValue = iTitleIndex[ikey];

      //先把標題送去取得對應value
      let outputTitle = comparisonTableList[ikey];

      //判斷
      switch (sortValueType(outputTitle)) {
        case "date":
          for (const j of outputTitle) {
            tempData[outputTableTitle.indexOf(j)] = modifyMailData(ikey, iData[iValue])[j];
          }
          break
        case "string":
          tempData[outputTableTitle.indexOf(outputTitle)] = modifyMailData(ikey, iData[iValue]);
          break
        case "error":
          throw new Error("data type not found.")
      }

    }

    //其餘信件資料寫入
    tempData[outputTableTitle.indexOf('mailFrom')] = iFrom;
    tempData[outputTableTitle.indexOf('mailAttFileName')] = iName;
    tempData[outputTableTitle.indexOf('mailSubject')] = iSubject;
    tempData[outputTableTitle.indexOf('mailDate')] = iDate;
    output.push(tempData);
  }
}

//取得標題對應的位置，並放入map供查詢 : {"upload": 2}
function setTitleIndex(iTitle, iData) {
  let iTitleMap = {};
  for (const str of iTitle) {
    iTitleMap[str] = iData.indexOf(str);
  }
  return iTitleMap;
}

//判斷map物件裡的value是不是array
function sortValueType(checkValue) {
  if ((typeof checkValue === 'object') && Array.isArray(checkValue)) {
    return 'date'
  }
  else if (typeof checkValue === 'string') {
    return 'string'
  }
  else {
    return 'error'
  }
}

function selectSheetTilte(fileName) {
  //各個附加檔案裡，要擷取的指定標題
  let openSignalTitle = ['dl_speed', 'ul_speed', 'my_lat', 'my_lon', 'CID', 'timestamp'];
  let ver1SpeedtestTitle = ['￦ﾗﾥ￦ﾜﾟ', '￧ﾷﾯ￥ﾺﾦ', '￧ﾶﾓ￥ﾺﾦ', '￤ﾸﾋ￨ﾼﾉ', '￤ﾸﾊ￥ﾂﾳ'];
  let ver2SpeedtestTitle = ['日期', '緯度', '經度', '下載', '上載'];
  let ver3SpeedtestTitle = ['日期', '緯度', '經度', '下載速度', '上傳速度'];

  let switchCase = 0;
  if (fileName.indexOf('osm_speedtests_') === 0) { switchCase = "osVer1"; }
  if (fileName.indexOf('SpeedtestPutMailBodyData') === 0) { switchCase = "spVer1"; }
  if (fileName.indexOf('SpeedTestExport_') === 0) { switchCase = "spVer2"; }
  if (fileName.indexOf('results.csv') === 0) { switchCase = "spVer3"; }

  //依結果回傳對應tiltle
  switch (switchCase) {
    case "osVer1":
      return openSignalTitle;
    case "spVer1":
      return ver1SpeedtestTitle;
    case "spVer2":
      return ver2SpeedtestTitle;
    case "spVer3":
      return ver3SpeedtestTitle;
    default:
      throw new Error("FileName not found.");
  }
}
