//let sender = '';
//sender is equal to the email address of the intended sender
let subject = '1-Nhjttm4FIH8x9HvDH79Ej-nocx1q-yjlBipYyCctD8';
//subject is equal to a string that is contained in the intended subject
let spreadSheetId = 'Banton';
//spreadSheetId is equal to the value of the spreadsheet from the URL

function getGmailEmails(){
var threads = GmailApp.getInboxThreads();
for(var i = 0; i < threads.length; i++){
		var messages = threads[i].getMessages();
		var msgCount = threads[i].getMessageCount();
	for (var j= 0; j <messages.length; j++){
		message = messages[j];
		if (message.isInInbox()){
			extractDetails(message,msgCount);
			}
		}
	}
}
function extractDetails(message, msgCount){
	var sheetname = 'Sheet1';
	var ss = SpreadsheetApp.openById(spreadSheetId);
	var timezone = 
SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
	var sheet = ss.getSheetByName(sheetname);

var dateTime = Utilities.formatDate(message.getDate(), timezone, "dd-MM-yy");
			var subjectText = message.getSubject();
			var fromSend = message.getFrom();
			var toSend = message.getTo();
			var bodyContent = message.getPlainBody();
	if (subjectText.includes(subject)){
		sheet.appendRow([dateTime, msgCount, fromSend, toSend, subjectText, 
bodyContent]);}
}
function onOpen() {
SpreadSheetApp.getUi()
.createMenu('Click to Fetch Emails')
.addItems('Get Email', 'getGmailEmails')
.addToUi();
}