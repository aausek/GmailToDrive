// Scan Gmail app for messages from specific sender email
function GmailScan() {

    // Gets latest message from sender email
    var threads = GmailApp.search('in:inbox from:"do-not-reply@candidatecare.com"');
    var messages = threads[0].getMessages();
    var message = messages[messages.length - 1];
    var msgDate = message.getDate();
    var attachments = message.getAttachments();
    for (var i = 0; i < attachments.length; i++) {
        if (attachments[i].getName() === 'candidates.xlsx') {
            var attachment = attachments[i];
        }
    }

    // Declare attachment name as var
    var filename = attachment.getName();
    // Does the attachment name match?
    if (filename === 'candidates.xlsx') {
        // Test file date and name with console.log
        console.log(msgDate);
        console.log(filename);

        // Call function to create google sheet with file name and date
        convertExceltoGS(filename, msgDate);
    }
}

// Convert excel file attachment to google sheet
function convertExceltoGS(filename, msgDate) {

    // From Google Drive App, search for file by name
    var excelFile = DriveApp.getFilesByName(filename).next();
    // From main Drive folder by ID
    var folderId = "1f_bAl6UoZr5V5W1D222GTvhzWkP5DZYO";
    var blob = excelFile.getBlob();

    // Create file with name and formatted date
    var file = {
        title: "candidates" + "_" + Utilities.formatDate(msgDate, "EST", 'yyyy-MM-dd'),
        parents: [
            {
                //"kind": "drive#parentReference",
                "id": folderId
            }
        ]
    };

    // Insert newly created google sheet into main Drive folder by ID above
    file = Drive.Files.insert(file, blob, {
        convert: true
    });
    console.log('converted file name: ', file.title);
    // Call function to copy contents of new google sheet file to a master copy
    CopyRange(msgDate);
}

// Function to copy contents of one google sheet into another
function CopyRange(msgDate) {
    // Open the file
    var FileNameString = 'candidates_' + Utilities.formatDate(msgDate, "EST", 'yyyy-MM-dd');
    var FileIterator = DriveApp.getFilesByName(FileNameString);
    // While loop to look for file by name
    while (FileIterator.hasNext()) {
        var file = FileIterator.next();
        if (file.getName() == FileNameString) {
            // After locating file, get its ID
            var fileID = file.getId();
        }
    }

    // Open located file by its ID
    var sss = SpreadsheetApp.openById(fileID); //replace with source ID
    // From its (single) sheet, copy contents
    var ss = sss.getSheetByName('New Hire Report (IT)'); //replace with source Sheet tab name
    var range = ss.getRange('A2:AL100'); //assign the range you want to copy
    var data = range.getValues();

    // Open master google sheet ('New_Hire') by ID and paste contents from previous file
    var tss = SpreadsheetApp.openById('1A6s5HZH0T26xhtkrBxzXsicpfNWvK2sPlYBUZKX-NEM'); //replace with destination ID
    // Specify which master google sheet to copy into
    var ts = tss.getSheetByName('Sheet1'); //replace with destination Sheet tab name
    // Ensure contents are pasted into next available blank row -- TO PREVENT OVERWRITING CONTENTS
    ts.getRange(ts.getLastRow() + 1, 2, 99, 38).setValues(data); // Define dimensions of copied data (row, column, numRows, numCols)
    console.log("Copied into master");

}
