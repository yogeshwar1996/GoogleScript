function onOpen() {
  var menu = [{name: 'Generate Report', functionName: 'generateReport'}];
  SpreadsheetApp.getActive().addMenu('Report', menu);
}


function generateReport() {
  
  var ss = SpreadsheetApp.getActive();
  // Mention Sheet
  var sheet = ss.getSheetByName('Feb');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var bill_url
  var table = [[]]
  // Title of Report 
  var doc = DocumentApp.create('Report for Expenses Sample Feb')
  var body = doc.getBody();
  var text = body.editAsText();

  body.insertParagraph(0, doc.getName()).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendPageBreak();
  for (var i = 1; i < values.length;i++) {
    var table_ =[['Field','Values']]
    for (var j = 0; j < values[i].length; j++) {    
       var row_ = []
       row_[0] = values[0][j]
       row_[1] = values[i][j];
       table_[j] = row_ ;
      if(j==5){
        bill_url = values[i][j];
      }
    }
    data = getBillCopy(bill_url)
    if(data.type != 'application/pdf' ){
      appendImageToBody(data.fileBlob, body)
    }
    else{
      text.appendText('\nImage not available. The attachment is a pdf.\n');
    } 
    body.appendTable(table_);
    body.appendPageBreak();
  }
  doc.saveAndClose();
}

function getBillCopy(bill_url){
  var file_id = bill_url.substring(33,66)

  var fileBlob = DriveApp.getFileById(file_id).getBlob();  
  var resource = {
        title: fileBlob.getName(),
        mimeType: fileBlob.getContentType()
  }
  Logger.log(resource.mimeType)
  
  return {fileBlob: fileBlob, type: resource.mimeType}
}

function appendImageToBody(blob, body){
  var image =  body.appendImage(blob)
  var height = image.getHeight()                             
  var width = image.getWidth()
  var ratio  = width/height
   Logger.log(ratio)
  
  if(width>640){
    Logger.log("When width is greater than 640")
    var newW = 640;
    var newH = parseInt(newW/ratio);
    Logger.log(newH)
    if (newH >400){
      newH = 400
      newW = ratio * newH
    }
     image.setWidth(newW).setHeight(newH)
  }
  else if (height >600){
        Logger.log("When height is greater than 600")

     var newH = 400;
     image.setWidth(width).setHeight(newH)
  }
}
