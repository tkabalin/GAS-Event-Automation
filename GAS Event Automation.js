// @ts-nocheck

// Developed by Thomas Kabalin

function onOpen() {
 	var ui = SpreadsheetApp.getUi();
 	ui.createMenu('Admin')
 	 	.addItem('Sort Events', 'sortSheet')
 	 	.addItem('Archive Old Events','archiveEvent')
 	 	.addSeparator()
 	 	.addItem('Notify...','notifyOther')
 	 	.addToUi();

}

function sortSheet() {
 	var range = SpreadsheetApp
 	 	 	 	 	.getActive()
 	 	 	 	 	.getSheetByName("Booking Form")
 	 	 	 	 	.getRange("A4:K400")

 	range.sort({column: 2, ascending: true});
 	range.sort({column: 1, ascending: true});

 	var range = SpreadsheetApp
 	 	 	 	 	.getActive()
 	 	 	 	 	.getSheetByName("Archive (2022)")
 	 	 	 	 	.getRange("A4:K400")

 	range.sort({column: 2, ascending: false});
 	range.sort({column: 1, ascending: false});
 	}


function notifyOther() {
 	clear();
 	 	// Get the message to be sent
 	 	var value = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
 	
 	 	// Get the notifications sheet 	
 	 	var sheet = SpreadsheetApp.getActiveSpreadsheet();
 	 	SpreadsheetApp.setActiveSheet(sheet.getSheetByName('Notifications'));

 	 	// Set the email addresses
 	 	var email = Browser.inputBox('Enter email address');

 	 	SpreadsheetApp.getActiveSheet().getRange('A2').setValue(email);

 	 	// Set the messaage
 	 	SpreadsheetApp.getActiveSheet().getRange(2,2).setValue(value);
 	 	hideSheet();

 	 	let numRows = 1; // Number of rows to process

 	 	// Get the range containing the email addresses and messages
 	 	const dataRange = SpreadsheetApp
 	 	 	 	 	 	 	 	.getActive()
 	 	 	 	 	 	 	 	.getSheetByName("Notifications")
 	 	 	 	 	 	 	 	.getRange(2,1,numRows,2)
 	 	// Fetch values for each row in the Range.
 	 	const data = dataRange.getValues();
 	 	for (let row of data) {
 	 	 	const emailAddress = row[0]; // First column
 	 	 	const message = row[1]; // Second column
 	 	 	let subject = 'Notification of Event';
 	 	 	//Send emails to emailAddresses which are presents in First column
 	 	 	MailApp.sendEmail(emailAddress, subject, message);
 	 	}
 	clear();

}


function clear() {
 	var range = SpreadsheetApp
 	 	 	 	 .getActive()
 	 	 	 	 .getSheetByName("Notifications")
 	 	 	 	 .getRange('A2:B7');
 	 	range.clearContent();
}

function hideSheet() {
 	var sheet = SpreadsheetApp
 	 	 	 	 	.getActive()
 	 	 	 	 	.getSheetByName("Notifications");
 	sheet.hideSheet();

}

function archiveEvent() {
 	var sheet = SpreadsheetApp.getActiveSpreadsheet();
 	SpreadsheetApp.setActiveSheet(sheet.getSheetByName('Booking Form')); 

 	var date = sheet.getRange('A4').getValue();
 	
 	var temp = new Date();
 	var dateToday = temp.getDate()+'/'+(temp.getMonth()+1)+'/'+temp.getFullYear();

 	var numBooked = SpreadsheetApp.getActiveSheet().getRange('L1').getValue(); 	


for (let i = 0; i < numBooked; i++) {
 	if(date < temp){ 	
 	// Get the event
 	var value = SpreadsheetApp
 	 	 	 	 	.getActive()
 	 	 	 	 	.getSheetByName("Booking Form")
 	 	 	 	 	.getRange('A4:K4')
 	 	 	 	 	.getValues(); 

 	// Get the Archive sheet
 	var sheet = SpreadsheetApp.getActiveSpreadsheet();
 	SpreadsheetApp.setActiveSheet(sheet.getSheetByName('Archive (2023)')); 
 	
 	//Get the number of events 
 	var numEvents = SpreadsheetApp.getActiveSheet().getRange('H1:K1').getValue();

 	// Insert the event into the archive sheet
 	SpreadsheetApp.getActiveSheet().getRange(numEvents+4,1,1,11).setValues(value);

 	// Delete the event from the booking form
 	var event = SpreadsheetApp
 	 	 	 	 	.getActive()
 	 	 	 	 	.getSheetByName("Booking Form");
 	 	event.deleteRow(4);
 
 	var date = SpreadsheetApp
 	 	 	 	 	.getActive()
 	 	 	 	 	.getSheetByName("Booking Form")
 	 	 	 	 	.getRange('A4')
 	 	 	 	 	.getValue();
 	 	 	 	 	
 	}
}
}