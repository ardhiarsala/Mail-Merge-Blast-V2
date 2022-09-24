/**

Mail Merge Blast V2 by Ardhi Arsala Rahmani

Developed as a modified version to the Google Apps Script Mail Merge Sample App at 
https://developers.google.com/apps-script/samples/automations/mail-merge

Includes added features for personalized attachment, cc, bcc, sender name, and mail subject customization 
directly on the Google Sheets ('parameter' sheet) as well as beautiful body-emails based on Google Docs formatting.

*/

//Creates the menu item Mail Blast and User Info.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Blast')
      .addItem('Get Template Draft', 'blastEmails')
      .addItem('User Info', 'uInfo')
      .addToUi();

}

//Creates the User Info Modal (you can ignore this part).
function uInfo(){
  var userInfo = HtmlService.createHtmlOutputFromFile('userguide').setWidth(450).setHeight(290);
  SpreadsheetApp.getUi().showModalDialog(userInfo,"Mail Merge Blast V2 (Version 2.02)");
}

//Opens a user input prompt for the GDocs url as the email body template.
function blastEmails(templateURL, sheet=SpreadsheetApp.getActiveSheet()) {
  if (!templateURL){
    templateURL = Browser.inputBox("Mail Blast", 
                                      "Paste the Google Docs URL of the Email Template:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (templateURL === "cancel" || templateURL == ""){ 
    return;
    }
  } 

//Checks for the 'parameter' sheet, if not found throws an error.
try{
const paramSheet = SpreadsheetApp.getActive().getSheetByName('parameter').getDataRange();
} catch(error){
  var popupError = HtmlService.createHtmlOutput("<link rel='stylesheet' href='https://ssl.gstatic.com/docs/script/css/add-ons1.css'><p>The sheet parameter is not found. Make sure to read the user guide accessible at my <a onclick=\"window.open(\'https://github.com/ardhiarsala/Mail-Merge-Blast-V2')\">Github</a></p><br>Accidentally deleted the parameter sheet? Get a new copy <a onclick=\"window.open(\'https://docs.google.com/spreadsheets/d/1-ia3zmP5qjtN8YTGbT7ilKQYPlW1qloO4ShR3PahiHg/copy\')\">here</a>").setWidth(400).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(popupError,"Parameter Error");
}

//Calls the active spreadsheet for mail merge blast execution.
var ssheads = SpreadsheetApp.getActiveSheet().getDataRange().getDisplayValues();
var ssrows = ssheads.slice(2);

//Calls the parameter spreadsheet for added mail data
const subjectCell = SpreadsheetApp.getActive().getSheetByName('parameter').getRange("B2").getDisplayValue();
const ccCell = SpreadsheetApp.getActive().getSheetByName('parameter').getRange("B3").getDisplayValue();
const bccCell = SpreadsheetApp.getActive().getSheetByName('parameter').getRange("B4").getDisplayValue();
const senderCell = SpreadsheetApp.getActive().getSheetByName('parameter').getRange("B5").getDisplayValue();

//Opens the GDocs by URL and gets its ID to mail merge based on active spreadsheet data. 
var docTemplate = DocumentApp.openByUrl(templateURL)
var template = DriveApp.getFileById(docTemplate.getId());

ssrows.forEach((row,index)=>{
  //forEach loop thrown to copy the GDocs template, change the values based on <<var>> established and converts to HTML for the email.
  if(row[1] == ''){
  var draft = template.makeCopy();
  var draftid = draft.getId() ;
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+draftid+"&exportFormat=html"
  var param = {
method      : "get",
headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
muteHttpExceptions:true,
  };
  var draftdoc = DocumentApp.openById(draftid);
  var draftbody = draftdoc.getBody();
  ssheads[1].forEach((heading,i)=>{

    const headers = heading.toUpperCase();
    draftbody.replaceText(`<<${headers}>>`,row[i])

  })

  draftdoc.saveAndClose();

  var html = UrlFetchApp.fetch(url,param);
  var email = row[0];
  GmailApp.sendEmail(email, subjectCell,"", {
    htmlBody:html,
    cc:ccCell,
    bcc:bccCell,
    name:senderCell
    });

/** Deprecated in Version 2.02 after bouncebacked responses reported in mass email blasts due to Google security measures 
  if(row[2].length >= 1){
    //If statement on whether there are attachment values. Converts Google Drive Documents (.doc, .docx), Presentations (Google Slides, .pptx) and PDF files into blob attachments if valid ID values exists, otherwise no attachments are included.
  var html = UrlFetchApp.fetch(url,param);
  var email = row[0];
  var links = "https://drive.google.com/uc?export=view&id="+row[2];
  var attachment = [UrlFetchApp.fetch(links).getBlob().setName(`Lampiran - ${row[0]} `+new Date())]
  GmailApp.sendEmail(email, subjectCell,"", {
    htmlBody:html, 
    attachments:attachment,
    cc:ccCell,
    bcc:bccCell,
    name:senderCell
    });
  } else{
  var html = UrlFetchApp.fetch(url,param);
  var email = row[0];
  GmailApp.sendEmail(email, subjectCell,"", {
    htmlBody:html,
    cc:ccCell,
    bcc:bccCell,
    name:senderCell
    });
  }
 **/

  //Inputs the sent status values on the 'STATUS' column
  var ssStatus = sheet.getRange(index+3,2,1,1);
  ssStatus.setValue('Sent at '+new Date())

  draft.setTrashed(true)

  }


})
}
