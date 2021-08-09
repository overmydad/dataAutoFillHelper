/*
=======================
os與sp要擷取的標題
=======================
*/
var openSignalTitle = ['dl_speed', 'ul_speed', 'my_lat', 'my_lon', 'CID', 'timestamp'];
var speedTestTitle = ['日期', '緯度', '經度', '下載速度', '上傳速度'];

/*
=======================
os與sp與輸出表單的標題對應表 (key&value)
=======================
*/
var comparisonTableList = {
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
}

/*
=======================
輸出表單的標題
=======================
*/
var combineTableTitle = ['Date', 'Time', 'DL', 'UL', 'Latitude', 'Longitude', 'EnbID&CellID', 'mailForm', 'mailAttFilename', 'mailSubject', 'mailDate'];

/*
=======================
gmail query serach (以key&value形式儲存資料)
https://support.google.com/mail/answer/7190?hl=zh-Hant
=======================
*/
var gmailQuery = {
  //'from':'',
  //'cc':'',
  //'after': Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd")
  //'after': "2018/04/15",
  'subject': 'Speedtest 結果（CSV 匯出）'
};