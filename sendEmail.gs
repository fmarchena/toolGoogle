var timeSleep= 10000;
var messageCheck = 'Delivery Status Notification (Failure)';

function recorrerSheetCell()
{
  
  Logger.log("INIT recorrerSheetCell");  
  var gSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = gSpreadSheet.getSheets()[0];
  var emailsColumn = sheet.getRange('A2:A').getValues();
  var subjectColumn = sheet.getRange('B2:B').getValues();
  var fileColumn = sheet.getRange('D2:D').getValues();
  var subjetMailCheck=[];
  var cant = emailsColumn.length; 
  var row=0; 
  var cellCheck ='';
  Logger.log(" Length :["+cant+"]");  
  var gmailAppObjectArray = [];
  for(row; row<=cant;row++)
  {
    var email= emailsColumn[row];
   
    if(email!="" & email!=undefined )
    {
   cellCheck = 'C' +(row+2)   
   var now = new Date();   
   var subject= subjectColumn[row];
    var file= fileColumn[row];
    var gmailAppObject= {};
    gmailAppObject.email= email;
    gmailAppObject.subject= subject +'-'+ now ;
     gmailAppObject.body =  subject  + now;  
    gmailAppObject.file= file;
    Logger.log( gmailAppObject.subject[row]); 
    gmailAppObjectArray[row] = gmailAppObject;
      sendEmail(gmailAppObject);
    Utilities.sleep(timeSleep);
    var checkEmail  =  checkSendEmailSubject(gmailAppObject.subject);
   var cellCheck = sheet.getRange(cellCheck);
    if (checkEmail==1)
    {
      cellCheck.setValue('Error: ' +messageCheck );
    }else
    {
      cellCheck.setValue('ok');
    }
    }
  }
  Logger.log(" Length :["+cant+"]"); 
  
  
 Logger.log("END recorrerSheetCell");  
  
}


function sendEmail(gmailAppObject) {
    Logger.log("INIT sendEmail"); 
  try{
    
    GmailApp.sendEmail(gmailAppObject.email[0],  gmailAppObject.subject, gmailAppObject.body)
    Logger.log("Email:" +gmailAppObject.email[0]+"subject :" +gmailAppObject.subject+ "Body:"+ gmailAppObject.body);  
  
  }catch(e)
  {
     Logger.log(" Error  in sendEmail  :["+e+"]");  
  }
    Logger.log("END  sendEmail"); 
  
}


function checkSendEmailSubject(subject)
{
  Logger.log("INIT checkSendEmailSubject");  
  var queryMail = 'subject:"'+subject+'"';
  var threads = GmailApp.search(queryMail);
  var messages = GmailApp.getMessagesForThreads(threads);
  
  for (var i=0; i< messages.length; i++)
  {
    for (var j=0; j< messages[i].length; j++)
    {
    var messageSubject= messages[i][j].getSubject();
      if(messageSubject == messageCheck)
      {
        return 1;
      }
    
     }// end second for
  }// end first for 
  Logger.log("END checkSendEmailSubject");   
  return 0;  
  
}



function main()
{
  recorrerSheetCell();
}
