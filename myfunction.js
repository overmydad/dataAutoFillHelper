/*
=======================
將原始csv檔案提取需要的資料，轉換到另外一個二維陣列上
=======================
*/
var myCsv = function(csvDataArray, csvTitleArray, csvMailINFO) {
  var outputData = [];
  for (var x = 0; x < csvDataArray.length; x++) {
    var csvData = Utilities.parseCsv(csvDataArray[x]);
    var csvTitle = csvTitleArray[x];
    var mailINFO = csvMailINFO[x];
    //儲存標題的indexOf位置
    var titleIndex = {};
    for (var i = 0; i < csvData.length; i++) {
      //先解析csv標題，取得所需資料的所在陣列位置，而後再取其餘列數的資料
      if (i > 0) {
        //要建構二維陣列之前，得先塞入一個一維陣列
        outputData.push([]);
        for (var k in titleIndex) {
          //要取出的標題與標題index數值 EX : ['timestamp','2']
          var tIkey = k;
          var tIvalue = titleIndex[k];
          //原始要取出的資料&已轉換要輸出的資料
          var orgData = csvData[i][tIvalue];
          var opData = setData(tIkey, orgData);
          //原csv標題轉換到輸出表格的標題 (可用陣列表示)
          var getopTit = comparisonTableList[tIkey];

          //處理輸出表格的標題的時間日期陣列
          if (typeof getopTit === 'object' && Array.isArray(getopTit)) {
            for (var y in getopTit) {
              //getopTit的第幾個標題
              //Logger.log(getopTit[y]);
              //getopTit[y]的標題在combineTableTitle的第幾個位置
              var getODTilIndex = combineTableTitle.indexOf(getopTit[y]);
              outputData[i - 1][getODTilIndex] = opData[y];
            }
          }
          //處理輸出表格的標題的一般格式
          else if (typeof getopTit === 'string') {
            var getODTilIndex = combineTableTitle.indexOf(getopTit);
            outputData[i - 1][getODTilIndex] = opData;
          }
        }
        //匯出其他信件資訊
        for (var k in mailINFO) {
          var index = combineTableTitle.indexOf(k);
          outputData[i - 1][index] = mailINFO[k];
        }
      }
      //讀取第一行標題，取出標題與標提對應位置的key&value物件
      else {
        for (var j = 0; j < csvTitle.length; j++) {
          var titleString = csvTitle[j];
          var titleIndexNum = csvData[i].indexOf(titleString);
          titleIndex[titleString] = titleIndexNum;
        }
      }
    }
  }
  return outputData;
}


/*
=======================
針對不同標題，處理個別資料
http://bomy.github.io/2017/03/26/avoid-the-switch-statement-in-Javascript/
=======================
*/
var setData = function(action, value) {
  var cidToEnbID_CellID = function(value) {
    const hex = 256;
    var enbID = (Math.floor(value / hex)).toString();
    var cellID = (value % hex).toString();
    return enbID + "/" + cellID;
  };
  var UTCDate = function(value) {
    var ret = new Date(parseInt(value));
    var retArray = [];
    retArray.push(Utilities.formatDate(ret, 'GMT', 'yyyy/MM/dd'));
    retArray.push(Utilities.formatDate(ret, 'GMT', 'HH:mm'));
    return retArray;
  }
  var strDate = function(value) {
    var ret = value.split(" ");
    var retArray = [];
    retArray.push(ret[0].replace(/-/gi, '/'));
    retArray.push(ret[1]);
    return retArray;
  }
  var bitToMbps_check = function(value) {
    if (value.toString().indexOf('.') > 0) {
      return value;
    } else {
      return value / 1000 / 1000 * 8;
    }
  }
  var bpsToMbps = function(value) {
    return value / 1000;
  }
  var editMethod = {
    //opensignal區域
    'dl_speed': bpsToMbps,
    'ul_speed': bpsToMbps,
    'CID': cidToEnbID_CellID,
    'timestamp': UTCDate,
    //speedtest區域
    '日期': strDate,
    '下載速度': bitToMbps_check,
    '上傳速度': bitToMbps_check
  }
  if (typeof editMethod[action] !== 'function') {
    return value;
  }
  return editMethod[action](value);
}

/*
=======================
自定物件key&values的toString方式
=======================
*/
var myObj = {
  toString: function(_obj) {
    var temp = [];
    for (var k in _obj) {
      if (_obj.hasOwnProperty(k)) {
        temp.push(k + ":" + _obj[k]);
      }
    }
    return temp;
  }
}
/*
=======================
創造命名為當下時間的工作表
http://www.alicekeeler.com/2016/04/10/google-apps-script-create-new-tabs/
http://www.alicekeeler.com/2015/11/06/google-apps-script-code-insert-a-sheet/
https://developers.google.com/apps-script/reference/spreadsheet/sheet
https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet
=======================
*/
function makeTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var templateSheet = ss.getActiveSheet();
  //var tz = ss.getSpreadsheetTimeZone();
  var date = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd_HH:mm');
  var worksheet = ss.getSheetByName(date);
  if (worksheet === null) {
    var len = ss.getSheets().length;
    //ss.insertSheet(date, len, {template: templateSheet});
    ss.insertSheet(date, len);
    //worksheet.clear();
  }
  return date;
}
/*
=======================
寫入myTable資料，到指定工作表
=======================
*/
function writeTabs(tabName, myTable) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var worksheet = ss.getSheetByName(tabName);
  //判斷表內有沒有資料，如果完全沒有，添加標題後再增加資料
  //如果有資料，直接安插資料在最後一排
  if (worksheet.getRange("A1").isBlank()) {
    worksheet.appendRow(combineTableTitle);
  }
  for (var x in myTable) {
    if (myTable[x].length > 0) {
      worksheet.appendRow(myTable[x]);
    }
  }
}
/*
=======================
製作自訂選單到試算表內，並連結函式
=======================
*/
function onOpen() {
  var menu = [{
    name: "讀取今天的mail資料",
    functionName: "mainFunction"
  }];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("擴充功能選單-Super App", menu);
}