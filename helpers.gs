function SendMail(){
MailApp.sendEmail("vinay.poulose@akanksha.org",
                    "Subject",
                    "\n\n" +
                    "Here are all members" + members,
                    {name:"From Name"});
}
function ReadData(){
    //var spreadsheetUrl = "https://drive.google.com/drive/folders/0B-R-0KdPxrZffldGd2R6QlBNcDBKa2stTXAxRkJyV2RSS3RoU283Y0VJTFBUM3NkTWUydlU";
  var spreadsheet = SpreadsheetApp.openById("1UQRGLs-YugdUDi3XB-1fCiLiuxLF4g1T");
  var sheets = spreadsheet.getSheets().sort();
  var members = sheets[0].getDataRange().getValues();
Logger.log(members);
  var temp=members;
}
function Delete(){
Drive.Files.remove("1yLIOjcOxTR82TdrhdwOjocKwsCHlAen3QqcBTfWaGhQ");
}
function Test(){
var fileName=Drive.Files.get("1UQRGLs-YugdUDi3XB-1fCiLiuxLF4g1T").title;
  var Anu=convertExceltoGoogleSpreadsheet(fileName);
  var Anu2=2;
}
function convertExceltoGoogleSpreadsheet(fileName) {
  
  try {
    
    // Written by Amit Agarwal
    // www.ctrlq.org

    fileName = fileName || "microsoft-excel.xlsx";
    
    var excelFile = DriveApp.getFilesByName(fileName).next();
    var fileId = excelFile.getId();
    var folderId = Drive.Files.get(fileId).parents[0].id;  
    var blob = excelFile.getBlob();
    var resource = {
      title: excelFile.getName(),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id: folderId}],
    };
    
    
    var test=Drive.Files.insert(resource, blob);
    var anu=test;
    
  } catch (f) {
    Logger.log(f.toString());
  }
  
}