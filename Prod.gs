// Globals
var folderId = "folder ID goes here"; // Google Drive "HR emails/candidates file" folder ID

// Scan Gmail app for messages from specific sender email
function ScanGmail() {

    // Logs information about any attachments in the first 100 inbox threads.
    //var threads = GmailApp.search('in:inbox from:"do-not-reply@candidatecare.com"', 0, 1);
    var threads = GmailApp.search('filename:name.extension', 0, 1);
    var msgs = GmailApp.getMessagesForThreads(threads);
    //var folderId = "1f_bAl6UoZr5V5W1D222GTvhzWkP5DZYO";
    var todayDate = Utilities.formatDate(new Date(), "EST", 'yyyy-MM-dd');
    for (var i = 0; i < msgs.length; i++) {
        for (var j = 0; j < msgs[i].length; j++) {
            var attachments = msgs[i][j].getAttachments();
            for (var k = 0; k < attachments.length; k++) {
                // Logger.log('Message "%s" contains the attachment "%s" (%s bytes)',
                //     msgs[i][j].getSubject(), attachments[k].getName(), attachments[k].getSize());
                var attachmentBlob = attachments[k].copyBlob();
                var attachmentName = attachments[k].getName();
                var msgDate = msgs[i][j].getDate();
                var attDate = Utilities.formatDate(msgDate, "EST", 'yyyy-MM-dd');
                //console.log("email date: ", attDate);
                //console.log("today date: ", todayDate);
                // Create file with name and formatted date
                if (todayDate == attDate) {
                    var file = {
                        title: attachmentName,
                        parents: [
                            {
                                //"kind": "drive#parentReference",
                                "id": folderId
                            }
                        ]
                    };
                    // // Insert newly created google sheet into main Drive folder by ID above
                    // file = Drive.Files.insert(file, attachmentBlob, {
                    //     convert: false
                    // });
                  var file = DriveApp.createFile(attachmentBlob);
                  var uploadFolder = DriveApp.getFolderById(folderId);
                  uploadFolder.addFile(file);
                  convertFile(attachmentName, attDate);
                  uploadFolder.removeFile(file);
                }
                else {
                  console.log("Nothing yet today!");
                  break;
                }
            }
        }
    }
}

// Convert excel file attachment to google sheet
function convertFile(attachmentName, attDate) {

    // From Google Drive App, search for file by name
    var excelFile = DriveApp.getFilesByName(attachmentName).next();
    // From main Drive folder by ID
    //var folderId = "1f_bAl6UoZr5V5W1D222GTvhzWkP5DZYO";
    var blob = excelFile.getBlob();

    // Create file with name and formatted date
    var file = {
        title: 'name_' + attDate,
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
    var fileName = file.title;
    console.log('converted file name: ', fileName);
    // Call function to copy contents of new google sheet file to a master copy
    CopyRange(fileName);
}

// Function to copy contents of one google sheet into another
function CopyRange(fileName) {
    // Open the file
    var FileNameString = fileName;
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
    var ss = sss.getSheetByName('Sheet name)'); //replace with source Sheet tab name
    var range = ss.getRange('A2:AL100'); //assign the range you want to copy
    var data = range.getValues();

    // Open master google sheet ('New_Hire') by ID and paste contents from previous file
    var tss = SpreadsheetApp.openById('master google sheet ID'); //replace with destination ID
    // Specify which master google sheet to copy into
    var ts = tss.getSheetByName('Sheet1'); //replace with destination Sheet tab name
    // Ensure contents are pasted into next available blank row -- TO PREVENT OVERWRITING CONTENTS
    ts.getRange(ts.getLastRow() + 1, 2, 99, 38).setValues(data); // Define dimensions of copied data (row, column, numRows, numCols)
    console.log("Copied into master");

}

// Sorting function - Gather requirements from Alex
