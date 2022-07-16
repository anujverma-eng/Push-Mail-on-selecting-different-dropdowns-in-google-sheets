var presenterSheetName = "Presenter";
var teamSheetName = "Team";

var emailQuotaRemaining = MailApp.getRemainingDailyQuota();

var ss = SpreadsheetApp.getActiveSheet();

var referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reference")


var emails = {
  "educator": referenceSheet.getRange(2, 3).getValue(),
  "sketchArtist": referenceSheet.getRange(3, 3).getValue(),
  "animator": referenceSheet.getRange(4, 3).getValue(),
  "pptDesigner": referenceSheet.getRange(5, 3).getValue(),
  "videoEditor": referenceSheet.getRange(6, 3).getValue(),
  "coordinator": referenceSheet.getRange(7, 3).getValue(),
  "other1": referenceSheet.getRange(8, 3).getValue(),
  "other2": referenceSheet.getRange(9, 3).getValue(),
  "other3": referenceSheet.getRange(10, 3).getValue(),
  "other4": referenceSheet.getRange(11, 3).getValue(),
  "other5": referenceSheet.getRange(12, 3).getValue(),
  "other6": referenceSheet.getRange(13, 3).getValue(),
  "other7": referenceSheet.getRange(14, 3).getValue(),
  "other8": referenceSheet.getRange(15, 3).getValue(),
  "other9": referenceSheet.getRange(16, 3).getValue(),
}

var names = {
  "educator": referenceSheet.getRange(2, 2).getValue(),
  "sketchArtist": referenceSheet.getRange(3, 2).getValue(),
  "animator": referenceSheet.getRange(4, 2).getValue(),
  "pptDesigner": referenceSheet.getRange(5, 2).getValue(),
  "videoEditor": referenceSheet.getRange(6, 2).getValue(),
  "coordinator": referenceSheet.getRange(7, 2).getValue(),
  "other1": referenceSheet.getRange(8, 2).getValue(),
  "other2": referenceSheet.getRange(9, 2).getValue(),
  "other3": referenceSheet.getRange(10, 2).getValue(),
  "other4": referenceSheet.getRange(11, 2).getValue(),
  "other5": referenceSheet.getRange(12, 2).getValue(),
  "other6": referenceSheet.getRange(13, 2).getValue(),
  "other7": referenceSheet.getRange(14, 2).getValue(),
  "other8": referenceSheet.getRange(15, 2).getValue(),
  "other9": referenceSheet.getRange(16, 2).getValue(),
}

function SendTheMail(Email, Subject, Body, options) {
  MailApp.sendEmail(Email, Subject, Body, options);
}

var inCC = `${emails.other1 && emails.other1},${emails.other2 && emails.other2},
${emails.other3 && emails.other3},${emails.other4 && emails.other4},
${emails.other5 && emails.other5},${emails.other6 && emails.other6},
${emails.other7 && emails.other7},${emails.other8 && emails.other8},${emails.other9 && emails.other9}, `;

function notificaitons() {
  var row = ss.getActiveCell().getRow();
  var col = ss.getActiveCell().getColumn();
  var playlistCode = ss.getRange(row, 1).getValue();

  if (ss.getSheetName().toString() === presenterSheetName) {
    switch (col) {
      case 3:

        if (ss.getActiveCell().getValue() === "Uploaded") {
          var subject = `${names.educator} Has Uploaded the Script`

          var emailBody = `Hello, Team \n${names.educator} has uploaded the Script and Audio for \nLecture ${playlistCode} 
          \nPFA the Link of Folder \n${ss.getRange(row, col - 1).getValue()} \n\nThanks and Regards \nK-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.coordinator };

          SendTheMail(`${emails.pptDesigner},${emails.sketchArtist},`, subject, emailBody, options);
        }
        break;
      case 6:
        if (ss.getActiveCell().getValue() === "Approved") {
          var subject = `PPT has been Approved`

          var emailBody = `Hello, Team \nPPT for Lecture Code:${playlistCode} has Approved \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.coordinator };

          SendTheMail(emails.pptDesigner, subject, emailBody, options);
        } else if (ss.getActiveCell().getValue() === "Changes Required") {
          var subject = `${names.educator}, Changes Required in PPT`

          var emailBody = `Hello, Team \nThere are some Changes Required in PPT for Lecture Code:${playlistCode}\n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.coordinator };

          SendTheMail(emails.pptDesigner, subject, emailBody, options);

          // add task for SME that changes are required.
          var rangeLink = getLinkToRange(`O${row}`)
          var objectArray = ["Changes Required in PPT", names.educator, getTime(), playlistCode, ss.getRange(row, col - 1).getValue(), names.pptDesigner, rangeLink];
          addTODO(objectArray);
          SpreadsheetApp.getActive().toast(`Task Given to SME for Changes in PPT \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 15);

        } else { }
        break;

      case 8:
        if (ss.getActiveCell().getValue() === "Uploaded") {
          var subject = `${names.educator}, has Uploaded the Lecture`

          var emailBody = `Hello, Team \n${names.educator} has Recorded the Lecture for Lecture Code:${playlistCode} \nPFA the Link to Recorded Lecture Folder \n${ss.getRange(row, col - 1).getValue()} \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.coordinator };

          SendTheMail(emails.videoEditor, subject, emailBody, options);

          // add task for Video Editor to start Basic Editing.
          var rangeLink = getLinkToRange(`S${row}`)
          var objectArray = ["Basic Editing", names.educator, getTime(), playlistCode, ss.getRange(row, col - 1).getValue(), names.videoEditor, rangeLink];
          addTODO(objectArray);
          SpreadsheetApp.getActive().toast(`Task Given to Video Editor to start Basic Editing for \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 15);

        }
        break;
      case 10:
        if (ss.getActiveCell().getValue() === "Done") {
          var subject = `${names.educator}, has Added the Timestamping Sheet`

          var emailBody = `Hello,${names.videoEditor} \n${names.educator} has Uploaded the Sheet for TimeStamp for Lecture Code:${playlistCode} \nPFA the Link to Timestamping sheet Folder \n${ss.getRange(row, col - 1).getValue()} \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.coordinator };

          SendTheMail(emails.videoEditor, subject, emailBody, options);

          // add task for video Editor to start content editing.
          var rangeLink = getLinkToRange(`V${row}`)
          var objectArray = ["Content Editing", names.educator, getTime(), playlistCode, ss.getRange(row, col - 1).getValue(), names.videoEditor, rangeLink];
          addTODO(objectArray);
          SpreadsheetApp.getActive().toast(`Task Given to Video Editor to start Content Editing for \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 15);

        }
        break;
      case 12:
        if (ss.getActiveCell().getValue() === "Final") {
          var subject = `Lecture is Final for, ${names.educator}`

          var emailBody = `Hello,${names.educator} \nThe Lectue is Final for Lecture Code:${playlistCode} \nPFA the Link to Final Lecture Folder \n${ss.getRange(row, col - 1).getValue()} \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator };

          SendTheMail(`${emails.other1},${emails.other2},${emails.videoEditor}`, subject, emailBody, options);
        }
        break;
      default:

    }
  } else if (ss.getSheetName().toString() === teamSheetName) {
    switch (col) {
      // case 3:
        // if (ss.getActiveCell().getValue() === "Approved") {
        //   var subject = `${names.educator}, your Script has Approved`

        //   var emailBody = `Hello, ${names.educator} \nYour Script for Lecture Code:${playlistCode} has Approved \n \nThanks and Regards \n K-10 Notifications`

        //   var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator + "," + emails.sketchArtist + "," + emails.pptDesigner + "," + emails.videoEditor + "," + emails.animator };

        //   SendTheMail(emails.educator, subject, emailBody, options);

        //   // as soon as Script gets approved, task for Sketch Artist.
        //   var rangeLink = getLinkToRange(`F${row}`)
        //   var objectArray = ["Create Sketches", names.educator, getTime(), playlistCode, `Script Link\n${ss.getRange(row, col - 1).getValue()}`, names.sketchArtist, rangeLink];
        //   addTODO(objectArray);
        //   SpreadsheetApp.getActive().toast(`Task Given to Sketch Artist to start Creating sketches for \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 10);
        // }
        // break;
      case 5:
        if (ss.getActiveCell().getValue() === "Done") {
          var subject = `Sketches task is done for: ${playlistCode}`

          var emailBody = `Hello,${names.educator} \nSketch Artist has uploaded the Sketches for \nLecture ${playlistCode} 
          \nPFA the Link of Sketches Folder \n${ss.getRange(row, col - 1).getValue()} \n\nThanks and Regards \nK-10 Notifications`

          var options = { cc: emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator };

          SendTheMail(`${emails.animator},${emails.pptDesigner},${emails.other1},${emails.other2}`, subject, emailBody, options);
          // as soon as Sketches task got done.
          var rangeLink = getLinkToRange(`O${row}`)
          var objectArray = ["Start Creating PPT", names.educator, getTime(), playlistCode, ss.getRange(row, col - 1).getValue(), `${names.pptDesigner}`, rangeLink];
          addTODO(objectArray);
          SpreadsheetApp.getActive().toast(`Task Given to PPT Designer to start creating PPT for \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 10);
          
        }
        break;

        case 6:
        if (ss.getActiveCell().getValue() === "Yes") {
          var subject = `Animations Required in : ${playlistCode}`

          var emailBody = `Hello,${names.animator} \nSketch Artist has uploaded the Sketches, Pls Start creating animations for \nLecture ${playlistCode} 
          \nPFA the Link of Sketches Folder \n${ss.getRange(row, col - 1).getValue()} \n\nThanks and Regards \nK-10 Notifications`

          var options = { cc: emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator };

          SendTheMail(`${emails.animator},${emails.pptDesigner},${emails.other1},${emails.other2}`, subject, emailBody, options);
          // as soon as Sketch artist said, that Animation is required.
          var rangeLink = getLinkToRange(`I${row}`)
          var objectArray = ["Start Creating Animations", names.educator, getTime(), playlistCode,ss.getRange(row, col - 1).getValue(), `${names.pptDesigner}`, rangeLink];
          addTODO(objectArray);
          SpreadsheetApp.getActive().toast(`Task Given to Animator to start creating Animations for \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 10);
          
        }
        break;
        
      case 9:
        if (ss.getActiveCell().getValue() === "Done") {
          var subject = `Animation task is done for: ${playlistCode}`

          var emailBody = `Hello,${names.educator} \nAnimator has Done his task and has uploaded the Animations for \nLecture ${playlistCode} 
          \nPFA the Link of Animations Folder \n${ss.getRange(row, col - 1).getValue()} \n\nThanks and Regards \nK-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator };

          SendTheMail(emails.videoEditor, subject, emailBody, options);
          // as soon as Animations task got done.
          var rangeLink = getLinkToRange(`P${row}`)
          var objectArray = ["Add Animations in PPT", names.educator, getTime(), playlistCode, `${ss.getRange(row, col - 5).getValue()}\nAnimations Folder Link \n${ss.getRange(row,col-1).getValue()}`, `${names.pptDesigner}`, rangeLink];
          addTODO(objectArray);
          SpreadsheetApp.getActive().toast(`Task Given to PPT Designer to start Adding animations in PPT for \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 10);
        }
        break;
      case 11:
        if (ss.getActiveCell().getValue() === "Given") {
          var subject = `SME has added the Input for Sketches : ${playlistCode}`

          var emailBody = `Hello,${names.sketchArtist} \nSME has added the Input for Sketches for \nLecture ${playlistCode} 
          \nPFA the Link of Input Folder \n${ss.getRange(row, col - 1).getValue()} \n\nThanks and Regards \nK-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator };

          SendTheMail(emails.sketchArtist, subject, emailBody, options);

          // as soon as SME Required any INPUT to sketchArtist.
          var rangeLink = getLinkToRange(`E${row}`)
          var objectArray = ["Create Sketches(Input added by PPT Designer)", names.educator, getTime(), playlistCode,ss.getRange(row, col - 1).getValue(), names.sketchArtist, rangeLink];
          addTODO(objectArray);
          SpreadsheetApp.getActive().toast(`Task Given to Sketch Artist for \nLecture Code: ${playlistCode} \nAt: ${getTime()}`, "Task Alloted", 10);
        }
        break;
      case 15:
        if (ss.getActiveCell().getValue() === "Round 2") {
          var subject = `${names.educator}, Pls Review PPT for Round 1`

          var emailBody = `Hello, ${names.educator} \nYour PPT for Lecture Code:${playlistCode} for Round 1 has Created Successfully \nPFA the Link to PPT Folder for Round 1 PPT \n${ss.getRange(row, col - 1).getValue()} \nKindly Review the PPT \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + "," + emails.other2 + "," + "," + emails.other3 + "," + "," + emails.other4 + "," + "," + emails.other5 + "," + "," + emails.coordinator };

          SendTheMail(emails.educator, subject, emailBody, options);
        } else if (ss.getActiveCell().getValue() === "Round 3") {
          var subject = `${names.educator}, Changes Done in PPT`

          var emailBody = `Hello,${names.educator} \nThe Required Changes in PPT for Lecture Code:${playlistCode} has Done. \nKindly Review the PPT and change the Status Accordingly \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + "," + emails.other2 + "," + "," + emails.other3 + "," + "," + emails.other4 + "," + "," + emails.other5 + "," + "," + emails.coordinator };

          SendTheMail(emails.educator, subject, emailBody, options);
        } else {
          var subject = `K-10, ${names.educator}, PPT is Ready for Review - Round 1`

          var emailBody = `Hello, ${names.other6} Sir \nThe K-10 PPT is Ready to Review for Lecture Code:${playlistCode} \nPFA the Link to PPT Folder for Round 1 PPT \n${ss.getRange(row, col - 1).getValue()} \nKindly Review the PPT \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator +","+ emails.educator +','+emails.pptDesigner};

          SendTheMail(emails.other6, subject, emailBody, options);
        }
        break;
      case 19:
        if (ss.getActiveCell().getValue() === "Done") {
          var subject = `${names.educator}, Basic Editing Done`

          var emailBody = `Hello, Team \nBasic Editing for Lecture Code:${playlistCode} has done \nPFA the Link to Basic Editing Folder \n${ss.getRange(row, col - 1).getValue()} \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator };

          SendTheMail(emails.educator, subject, emailBody, options);
        }
        break;
      case 22:
        if (ss.getActiveCell().getValue() === "Done") {
          var subject = `${names.educator}, Content Editing Done`

          var emailBody = `Hello, Team \nContent Editing for Lecture Code:${playlistCode} has done \nPFA the Link to Content Editing Folder \n${ss.getRange(row, col - 1).getValue()} \n \nThanks and Regards \n K-10 Notifications`

          var options = { cc: emails.other1 + "," + emails.other2 + "," + emails.other3 + "," + emails.other4 + "," + emails.other5 + "," + emails.coordinator };

          SendTheMail(emails.educator, subject, emailBody, options);
        }
        break;
      default:

    }
  } else { }

}



