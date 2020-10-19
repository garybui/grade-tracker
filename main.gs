//Change these constants to match your information
var OWNER_ID = ''; //Use getOwnerId() function to find out your ownerID
var EMAIL_TEMPLATE_ID = "";
var IMPORT_GUARDIAN_EMAILS_FROM = ""; //leave blank for manual entry instead


//Change these variables to customize the spreadsheet 
var COURSE_DICTIONARY = [{key: "history", value: "history"}, {key: "science", value: "science"}];

var ASSIGNMENT_TYPE_DEFAULT = "OPTIONAL"; 

var ASSIGNMENT_TYPE1 = "HOMEWORK";
var TYPE1_KEYWORDS = ["homework", "assignment", "hw", "quiz"];

var ASSIGNMENT_TYPE2 = "CLASSWORK";
var TYPE2_KEYWORDS = ["classwork", "in class", "async", "cw", "exit ticket"];

var ASSIGNMENT_TYPE3 = "EXTRA CREDIT";
var TYPE3_KEYWORDS = ["extra credit"];

var DEFAULT_NUM_TOPICS = 3;

var EXCLUDED_CLASSES_KEYWORDS = ["advisory"];

var GRACE_PERIOD_HOURS = 1;  

var EMAIL_SUBJECT = "Missing Assignments for"; //Class name will be added later

/*---------------------------------------*/
// Be careful editing beyond this line //
/*---------------------------------------*/

var TAB_ONE_NAME = "Email Hub"
var TAB_TWO_NAME = "Master Roster"

var EMAIL_TEMPLATE_BREAK_IDENTIFIER = "[BREAK]";

var PROCEED = "yes";
var DECLINE = "no";
var END_FUNCTION = "end";

var CLASS_COLUMN_NAME = "Class";
var FN_COLUMN_NAME = "First Name";
var LN_COLUMN_NAME = "Last Name";
var EMAIL_COLUMN_NAME = "Email Address";
var GUARDIAN_ONE_COLUMN_NAME = "Guardian Email 1";
var GUARDIAN_TWO_COLUMN_NAME = "Guardian Email 2";
var DRAFT_ID_COLUMN_NAME = "Draft ID";
var GR_COLUMN_NAME = "GR Updated On";
var DATE_DRAFTED_COLUMN_NAME = "Date Drafted";
var DEADLINE_COLUMN_NAME = "Deadline";
var NUM_TOPICS_COLUMN_NAME = "Num Topics to View";
var EMAIL_HUB_HEADER = [CLASS_COLUMN_NAME, FN_COLUMN_NAME, LN_COLUMN_NAME, EMAIL_COLUMN_NAME, GUARDIAN_ONE_COLUMN_NAME, GUARDIAN_TWO_COLUMN_NAME, 
  DRAFT_ID_COLUMN_NAME, GR_COLUMN_NAME, DATE_DRAFTED_COLUMN_NAME, DEADLINE_COLUMN_NAME, DEADLINE_COLUMN_NAME, NUM_TOPICS_COLUMN_NAME];

var MISSING_ASSIGNMENT_LIST_DELIMITER = "****************\n";

var NUM_TOPICS_TO_VIEW = getNumTopicsToView();
var GRADE_REPORT = SpreadsheetApp.getActiveSpreadsheet();
var [EMAIL_HUB, MASTER_ROSTER, ...CLASS_LIST] = GRADE_REPORT.getSheets();

var CURRENT_COURSE_ID = "";
var CURRENT_COURSE_NAME = "";
var CURRENT_ASSIGNMENT_LIST = [];
var TOTAL_TYPE1 = [];
var TOTAL_TYPE2 = [];


function getOwnerId() {
  let courses = Classroom.Courses.list().courses;
  courses.forEach((course) => {Logger.log(course.ownerId, course.name)});
}

function getNumTopicsToView() {
  let emailHub = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_ONE_NAME);
  if(emailHub == null) {return DEFAULT_NUM_TOPICS;}
  else{
    let header = emailHub.getSheetValues(1,1,1,emailHub.getLastColumn()); 
    let targetValue = emailHub.getRange(2,getIndexOf(header, NUM_TOPICS_COLUMN_NAME)).getDisplayValue();
    (targetValue == "") ? targetValue = DEFAULT_NUM_TOPICS : targetValue = targetValue;
    return targetValue;
  }
}

function initializeSpreadsheet(e) {
  whenOpen(e);
  EMAIL_HUB = createEmailHubSheet();
  MASTER_ROSTER = createMasterRosterSheet();
}

function createEmailHubSheet() {
  let emailHub = GRADE_REPORT.getSheetByName(TAB_ONE_NAME);
  if(emailHub == null) {
    GRADE_REPORT.renameActiveSheet(TAB_ONE_NAME);
    emailHub = GRADE_REPORT.getSheetByName(TAB_ONE_NAME);
    let emailHubHeader = EMAIL_HUB_HEADER
    emailHub.appendRow(emailHubHeader);
    emailHub.getRange(2,emailHubHeader.indexOf(NUM_TOPICS_COLUMN_NAME)+1).setValue(DEFAULT_NUM_TOPICS);
  }
  
  return emailHub;
}

function createMasterRosterSheet() {
  let masterRoster = GRADE_REPORT.getSheetByName(TAB_TWO_NAME) || GRADE_REPORT.insertSheet(TAB_TWO_NAME);
  let masterRosterHeader = createMasterRosterHeader(NUM_TOPICS_TO_VIEW);
  masterRoster.clear();
  masterRoster.appendRow(masterRosterHeader);
  return masterRoster;
}

function createMasterRosterHeader(numOfTopics) {
  let masterRosterHeader = ["Class", "First Name", "Last Name", "Email Address"];
  for(i = 0; i < numOfTopics; i++){
    let unitNumber = String.fromCharCode('A'.charCodeAt(0) + i);
    let unitHW = "Unit ".concat(unitNumber, ": HW");
    let unitCW = "Unit ".concat(unitNumber, ": CW");
    let unitMW = "Unit ".concat(unitNumber, ": MW");
    masterRosterHeader.push(unitHW, unitCW, unitMW);
  }
  return masterRosterHeader;
}

function whenOpen(e) {
  let ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem("Control Panel", "controlPanel")
    .addToUi();           
}

function controlPanel() {
  let html = HtmlService.createHtmlOutputFromFile('controlpanel.html')
                        .setTitle("Control Panel");
  SpreadsheetApp.getUi().showSidebar(html);
}

function test()
{
  updateOnButtonClick("GR Updated On");
}

function updateOnButtonClick(targetItem) {
  let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn());
  let targetValue = EMAIL_HUB.getRange(2,getIndexOf(header, targetItem))
  if(targetItem == "GR Updated On") {createCourseRoster();}
  else if(targetItem == "Date Drafted") {
    if(!dueDateSpecified()){
      if(missingDueDateError()) generateEmailDrafts();
    }
    else if(!dueDateValid()){
      if(missingDueDateError()) generateEmailDrafts();
    }
    else generateEmailDrafts();
  }  
  else if(targetItem == "Deadline") {
    let newDeadline = promptForDueDate();
    if(newDeadline != "") targetValue.setValue(newDeadline);
  }
  else if(targetItem == "Num of Topics to View") updateNumOfTopicsToView();
  
  if(targetValue.getDisplayValue() == "") return "Has not been generated yet";
  else return targetValue.getDisplayValue();
}

function missingDueDateError() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert('Missing or Invalid Due Date', 'Would you like to proceed anyways?', ui.ButtonSet.YES_NO);
  return response == ui.Button.YES
}

function dueDateValid() {
  let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn());
  let dueDate = EMAIL_HUB.getRange(2,getIndexOf(header,DEADLINE_COLUMN_NAME)).getDisplayValue();
  let today = new Date();
  let month = today.getMonth()+1;
  let day = today.getDate();
  
  //User specified due date is in format "MM/DD (DAY OF THE WEEK)" 
  let dueDateMonth = dueDate.slice(0,dueDate.indexOf("/")); 
  let dueDateDay = dueDate.slice(dueDate.indexOf("/")+1,dueDate.indexOf("("));
  
  if(dueDateMonth > month && dueDateDay < day) return true;
  if(dueDateMonth >= month && dueDateDay > day) return true;
  return false;
}

function dueDateSpecified() {
  let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn());
  let dueDate = EMAIL_HUB.getRange(2,getIndexOf(header,DEADLINE_COLUMN_NAME)).getValue();
  
  return dueDate != "";
}

function promptForDueDate() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Update Due Date', 'When is the new due date? Ex: 7/26 (Sunday)', ui.ButtonSet.OK_CANCEL);
  let newDueDate = response.getResponseText();
  
  if (response.getSelectedButton() == ui.Button.OK) {return newDueDate;} 
}

function updateNumOfTopicsToView() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Change Number of Topics to View', 'How many topics would you like to view?', ui.ButtonSet.OK_CANCEL);
  let newNumTopics = response.getResponseText();
  
  if (response.getSelectedButton() == ui.Button.OK) {
    let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn()); 
    let targetValue = EMAIL_HUB.getRange(2,getIndexOf(header, NUM_TOPICS_COLUMN_NAME)).setValue(newNumTopics);
  }
}

function getLastUpdate(targetItem) {
  let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn());
  let targetValue = EMAIL_HUB.getRange(2,getIndexOf(header, targetItem))
  if(targetValue.getDisplayValue() == "") return "Has not been generated yet";
  else return targetValue.getDisplayValue();
}

function actionCompleteNotifier(status,close) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let htmlApp = HtmlService.createTemplateFromFile("popup.html");     
  htmlApp.data = status;
  htmlApp.close = close;
  ss.show(htmlApp.evaluate()
    .setWidth(300)
    .setHeight(200));
}

function getAllStudents(courseId) {
  let returnedStudents = []
  let options = {pageSize: 30};
    
  do {
    let search = Classroom.Courses.Students.list(courseId, options);
      
    if (search.students)
      Array.prototype.push.apply(returnedStudents, search.students);
      
    options.pageToken = search.nextPageToken;
  } while (options.pageToken);
    
  return returnedStudents;
}

function determineAssignmentType(assignmentName) {
  let assignmentType = ASSIGNMENT_TYPE_DEFAULT;
              
  for (let keyword of TYPE1_KEYWORDS) {
    if(assignmentName.toLowerCase().includes(keyword)) {
       assignmentType = ASSIGNMENT_TYPE1;
      break;
    }
  }          
  if(assignmentType == ASSIGNMENT_TYPE1) return assignmentType;    
  
  for (let keyword of TYPE2_KEYWORDS) {
    if(assignmentName.toLowerCase().includes(keyword)) {
      assignmentType = ASSIGNMENT_TYPE2;       
      break;
    }
  }          
  if(assignmentType == ASSIGNMENT_TYPE2) return assignmentType;
    
  for (let keyword of TYPE3_KEYWORDS) {
    if(assignmentName.toLowerCase().includes(keyword)) {
      assignmentType = ASSIGNMENT_TYPE3;
      break
    }
  }          
    
  return assignmentType;
}


function createCourseRoster() {
  actionCompleteNotifier("","");
  let rawCourses = getMyActiveCourses(OWNER_ID); 
  
  let classSizesPreviousIteration = [];
  let classSizesCurrentIteration = [];
  if(CLASS_LIST != []) CLASS_LIST.forEach(sheet => classSizesPreviousIteration.push(sheet.getLastRow()-4));
  

  createMasterRosterSheet(); 
  let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn());

  rawCourses.forEach(({name: courseName, id: courseId, section: courseSection}) => { 
    CURRENT_COURSE_ID = courseId;
    CURRENT_COURSE_NAME = courseName;
    let excludeClass = false;
    EXCLUDED_CLASSES_KEYWORDS.forEach(function(keyword) {
      if(courseName.toLowerCase().includes(keyword)) {
        excludeClass = true
        return;
      }
    });
    
    let rawStudents = getAllStudents(courseId);
    
    if(excludeClass || rawStudents == [] || rawStudents.length == 0) return;
    
    let tabName = courseName;
    if(courseSection != undefined) tabName = tabName.concat(" (", courseSection, ")"); 
    let sheet = GRADE_REPORT.getSheetByName(tabName) || GRADE_REPORT.insertSheet(tabName, SpreadsheetApp.getActiveSpreadsheet().getSheets().length);
    if(sheet.getLastRow()-4 > 0) classSizesPreviousIteration.push(sheet.getLastRow()-4);
    sheet.clear();

    
    let rawTopics = Classroom.Courses.Topics.list(courseId).topic || [];
    let rawAssignments = Classroom.Courses.CourseWork.list(courseId).courseWork || [];
    
    organizeRawAssignments(rawTopics, rawAssignments);
    
    let everyStudentAssignmentSubmission = getAllAssignmentSubmissions();
    
    let studentsWithAssignmentStatuses = processStudentSubmissions(rawStudents, everyStudentAssignmentSubmission);
   
    printTopicAndAssignmentHeaders(sheet, CURRENT_ASSIGNMENT_LIST);
    printStudents(sheet, studentsWithAssignmentStatuses);
  });
  CLASS_LIST.forEach(sheet => classSizesCurrentIteration.push(sheet.getLastRow()-4));
  manageEmailHubList(classSizesCurrentIteration, classSizesPreviousIteration);
  
  formatSheets();
  formatMasterRoster();
  formatEmailHub();
  
  EMAIL_HUB.getRange(2,getIndexOf(header, GR_COLUMN_NAME)).setValue(new Date());
  actionCompleteNotifier("Finished!","");
} 

function printStudents(sheet, students) {
  students.forEach(({profile, assignmentSubmissionStatus, missingType1Assignments, missingType2Assignments, missingType3Assignments}) => {
    let studentInfo = [profile.name.givenName, profile.name.familyName]
    if(studentInfo[0] != null || studentInfo[1] != null) {
      let printToClassTab = studentInfo.concat("", assignmentSubmissionStatus);
      sheet.appendRow(printToClassTab);

      
      let allMissingAssignmentsList = [];
      for (i = 0; i < missingType1Assignments.length; i++) {
        let numMissingType1 = missingType1Assignments[i].split("\t-").length-1;
        let numMissingType2 = missingType2Assignments[i].split("\t-").length-1;
        allMissingAssignmentsList.push(
          numMissingType1+"/"+TOTAL_TYPE1[i],
          numMissingType2+"/"+TOTAL_TYPE2[i],
          missingType1Assignments[i].concat(MISSING_ASSIGNMENT_LIST_DELIMITER, missingType2Assignments[i], MISSING_ASSIGNMENT_LIST_DELIMITER, missingType3Assignments[i])
        );
      }
      
      let printToMasterRoster = [CURRENT_COURSE_NAME].concat(studentInfo,profile.emailAddress, allMissingAssignmentsList)
      MASTER_ROSTER.appendRow(printToMasterRoster);
    }
  });
}


function processStudentSubmissions(rawStudents, allSubmissions) {
    rawStudents.forEach((student) => {
      let processedSubmissionStates = [];
      let missingTYPE1 = [];
      let missingTYPE2 = [];
      let missingTYPE3 = [];
      let thisStudentsSubmissions = allSubmissions.filter((assignmentSubmission) => assignmentSubmission.userId == student.userId || assignmentSubmission == "");
      thisStudentsSubmissions.forEach((submission) => {  
        if(typeof submission != "object") {
          missingTYPE1.push("");
          missingTYPE2.push("");
          missingTYPE3.push("");
          processedSubmissionStates.push("");
        } 
        else {
          let currentAssignment = CURRENT_ASSIGNMENT_LIST.find((assignment) => assignment.id == submission.courseWorkId);
          let {dueTime: assignmentDueTime, assignmentType: type, title: assignmentName} = currentAssignment;
                  switch(submission.state) {
                    case "NEW":
                    case "RECLAIMED_BY_STUDENT":
                    case "CREATED":
                      if(type == ASSIGNMENT_TYPE_DEFAULT) {
                        processedSubmissionStates.push("OPTIONAL");
                        break;
                      }
                      else if(type == ASSIGNMENT_TYPE3) {
                        processedSubmissionStates.push("EXTRA CREDIT");
                        missingTYPE3[missingTYPE3.length-1]= missingTYPE3[missingTYPE3.length-1].concat("\t-", assignmentName, "\n");
                        break;
                      }
                      if(!submission.late && assignmentDueTime != null) {
                          processedSubmissionStates.push("NOT YET DUE"); 
                          break;
                      }
                      else {
                          if(assignmentDueTime != null) {
                            processedSubmissionStates.push("MISSING");
                            type == ASSIGNMENT_TYPE1 ? 
                              missingTYPE1[missingTYPE1.length-1]= missingTYPE1[missingTYPE1.length-1].concat("\t-", assignmentName, "\n") : 
                              missingTYPE2[missingTYPE2.length-1]= missingTYPE2[missingTYPE2.length-1].concat("\t-", assignmentName,"\n"); 
                            break;
                          }
                          else processedSubmissionStates.push("NOT YET DUE");
                          break;
                      }
                    case "RETURNED":
                      processedSubmissionStates.push("GRADED");
                      break;
                    case "TURNED_IN":
                      if(submission.assignedGrade == null) {
                        if(assignmentDueTime == null) processedSubmissionStates.push("SUBMITTED");
                        else if(submission.late && !submittedWithinGracePeriod(submission.updateTime, assignmentDueTime)) {
                          let submissionDate = splitTimeStamp(submission.updateTime);
                          processedSubmissionStates.push(submissionDate.concat(" - LATE"));
                        }
                        else processedSubmissionStates.push("SUBMITTED"); 
                      }
                      else processedSubmissionStates.push("GRADED");
                      break;
                      
                    default:
                      processedSubmissionStates.push("ERROR");
                  }
            }
       });
       
       student.assignmentSubmissionStatus = processedSubmissionStates;
       student.missingType1Assignments = missingTYPE1;
       student.missingType2Assignments = missingTYPE2;
       student.missingType3Assignments = missingTYPE3;
    });

  return rawStudents;
}

function getAllAssignmentSubmissions() {
  let allAssignmentsSubmissions = [];
  CURRENT_ASSIGNMENT_LIST.forEach((assignment) => {
    if(typeof assignment == "object") {
      let assignmentSubmissions = Classroom.Courses.CourseWork.StudentSubmissions.list(CURRENT_COURSE_ID,assignment.id).studentSubmissions;
      allAssignmentsSubmissions = allAssignmentsSubmissions.concat(assignmentSubmissions);
    }
    else allAssignmentsSubmissions.push("");
  });
  
  return allAssignmentsSubmissions;
  
}

function printTopicAndAssignmentHeaders(sheet, assignmentsOrderedByTopic) {
    let printableTopics = ["","","Unit Name"];
    let printableAssignmentNames = ["","","Assignment Name"];
    let printableAssignmentDueDates = ["","","Due Date"];
    assignmentsOrderedByTopic.forEach((assignment) => {
      if(typeof assignment == "object") {
        let {title: assignmentName, dueDate: assignmentDueDate} = assignment;
        printableAssignmentNames.push(assignmentName);
        assignmentDueDate == undefined ? printableAssignmentDueDates.push("N/A") : printableAssignmentDueDates.push(assignmentDueDate.month+"/"+assignmentDueDate.day);
        printableTopics.push("");
      }
      else {
        printableTopics.push(assignment);
        printableAssignmentNames.push("");
        printableAssignmentDueDates.push("");
      }
    });

    sheet.appendRow(printableTopics);
    sheet.appendRow(printableAssignmentNames);
    sheet.appendRow(printableAssignmentDueDates);
    sheet.appendRow(["Student Names"]);

}

function organizeRawAssignments(rawTopics, rawAssignments) {
  let numValidTopics = 0;
  let assignmentsOrderedByTopic = [];
  
  rawTopics.forEach((topic) => {
    if(rawAssignments.length > 0) {
      let assignmentsFilteredByTopic = rawAssignments.filter((assignment) => assignment.topicId == topic.topicId);
      let orderedAssignmentsByCreationTime = assignmentsFilteredByTopic.sort((a,b) => (a.creationTime < b.creationTime) ? 1: -1);
      let processedAssignmentTypes = orderedAssignmentsByCreationTime
      processedAssignmentTypes.forEach((assignment) => assignment.assignmentType = determineAssignmentType(assignment.title));
      
      let numType1Assignments = processedAssignmentTypes.filter((assignment) => assignment.assignmentType == ASSIGNMENT_TYPE1).length;
      let numType2Assignments = processedAssignmentTypes.filter((assignment) => assignment.assignmentType == ASSIGNMENT_TYPE2).length;
      
      if(processedAssignmentTypes.length != 0 && numValidTopics < NUM_TOPICS_TO_VIEW) {
          numValidTopics++;          
          assignmentsOrderedByTopic.push(topic.name);
          assignmentsOrderedByTopic = assignmentsOrderedByTopic.concat(processedAssignmentTypes);
          TOTAL_TYPE1.push(numType1Assignments);
          TOTAL_TYPE2.push(numType2Assignments);
          
      }
    }
  });
  
  CURRENT_ASSIGNMENT_LIST = assignmentsOrderedByTopic;
}


function manageEmailHubList(classSizesCurrentIteration, classSizesPreviousIteration) {
   let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn());
   if(classSizesPreviousIteration == undefined || classSizesPreviousIteration < 1){
     let rangeToCopy = MASTER_ROSTER.getRange(2,1,MASTER_ROSTER.getLastRow(), 4);
     rangeToCopy.copyTo(EMAIL_HUB.getRange(2,1), {contentsOnly:true});
     
     let studentEmails = EMAIL_HUB.getRange(2,EMAIL_HUB_HEADER.indexOf(EMAIL_COLUMN_NAME)+1,EMAIL_HUB.getLastRow()-1,1).getValues().flat(1);
     
     studentEmails.forEach((email,matchIndex) => {
       let guardianEmails = getGuardianEmails(email);
       EMAIL_HUB.getRange(matchIndex +2, EMAIL_HUB_HEADER.indexOf(GUARDIAN_ONE_COLUMN_NAME)+1).setValue(guardianEmails[0]);
       EMAIL_HUB.getRange(matchIndex +2, EMAIL_HUB_HEADER.indexOf(GUARDIAN_TWO_COLUMN_NAME)+1).setValue(guardianEmails[1]);
     });
   }
   else {
     for(i = 0; i < classSizesCurrentIteration.length; i++) {
       if(classSizesCurrentIteration[i] - classSizesPreviousIteration[i] > 0){
         let startingPosition = 2;
         for(classSizeIterator = 0; classSizeIterator < i; classSizeIterator++) {
           startingPosition += classSizesPreviousIteration[classSizeIterator];
         } 
         
         let classRoster = MASTER_ROSTER.getRange(startingPosition,2,classSizesCurrentIteration[i], 1).getValues().flat(1);
         if(classSizesPreviousIteration[i] == 0) {
           let rangeToCopy = MASTER_ROSTER.getRange(startingPosition+1,1,classSizesCurrentIteration[i],4);
           rangeToCopy.copyTo(EMAIL_HUB.getRange(startingPosition+1,1), {contentsOnly:true});
         }
         else{       
           let emailHubRoster = EMAIL_HUB.getRange(startingPosition, 2, classSizesPreviousIteration[i], 1).getValues().flat(1);
           
           classRoster.forEach((student,index2) => {
             if(!emailHubRoster.includes(student)) {
                 let rangeToCopy = MASTER_ROSTER.getRange(startingPosition+index2,1,1,4);
                 EMAIL_HUB.insertRowAfter(startingPosition);
                 rangeToCopy.copyTo(EMAIL_HUB.getRange(startingPosition+1,1), {contentsOnly:true});
                 let studentEmail = EMAIL_HUB.getRange(startingPosition+1, getIndexOf(header, EMAIL_COLUMN_NAME)).getDisplayValue();
                 let guardianEmails = getGuardianEmails(studentEmail);
                 EMAIL_HUB.getRange(startingPosition+1, EMAIL_HUB_HEADER.indexOf(GUARDIAN_ONE_COLUMN_NAME)+1).setValue(guardianEmails[0]);
                 EMAIL_HUB.getRange(startingPosition+1, EMAIL_HUB_HEADER.indexOf(GUARDIAN_TWO_COLUMN_NAME)+1).setValue(guardianEmails[1]);
             }
           });
         }
       }
       if(classSizesCurrentIteration[i] - classSizesPreviousIteration[i] < 0){
         let startingPosition = 2;
         for(classSizeIterator = 0; classSizeIterator < i; classSizeIterator++){
           startingPosition += classSizesPreviousIteration[classSizeIterator];
         } 
         let classRoster = MASTER_ROSTER.getRange(startingPosition,2,classSizesCurrentIteration[i], 1).getValues().flat(1);
         let emailHubRoster = EMAIL_HUB.getRange(startingPosition, 2, classSizesPreviousIteration[i], 1).getValues().flat(1);
           
         emailHubRoster.forEach((student,index2) => {
           if(!classRoster.includes(student)){ 
               let rangeToDelete = EMAIL_HUB.getRange(startingPosition+index2,1,1,7);
               rangeToDelete.deleteCells(SpreadsheetApp.Dimension.ROWS);
           }
         });
       }
     }
   }
}

function getMyActiveCourses(ownerId) {
  let courses = Classroom.Courses.list().courses;
  return courses.filter((course) => {
     return course.courseState == "ACTIVE" && course.ownerId == ownerId;                              
  })
}

function mergeAndCenterCells(sheet,startingRow, startingColumn,numberOfRows, numberOfColumns) {
  if(numberOfColumns > 0) {
    let range = sheet.getRange(startingRow,startingColumn,numberOfRows,numberOfColumns)
    range.merge();
    range.setHorizontalAlignment("center");
    range.setVerticalAlignment("middle");
    range.setWrap(true);
  }
}

function getIndexOf(twoDimsArray, searchParam) {
  for (i = 0; i < twoDimsArray.length; i++) {
    var index = twoDimsArray[i].indexOf(searchParam);
    if (index > -1) {
      return index+1;
    }
  }
}

function splitTimeStamp(timeStamp) {
  let [removedTime, ...rest] = timeStamp.split("T");
  let delimitedByDash = removedTime.split("-");
  let [year, month, day] = delimitedByDash;
  let formattedDate = month + "/" + day;
  return formattedDate;
}

function submittedWithinGracePeriod(timeStamp, dueTime) {
  let submissionTime = timeStamp.slice(timeStamp.indexOf("T")+1);
  let submissionHour = submissionTime.slice(0,submissionTime.indexOf(":"));
  let submissionMinute = submissionTime.slice(submissionTime.indexOf(":")+1,submissionTime.indexOf(":")+3);
  
  let dueHour = dueTime.getHours();

  return (submissionHour == dueHour ||submissionHour == (dueHour + GRACE_PERIOD_HOURS)) 
}

function countNewStudents() {
  let emailHubRosterLength = EMAIL_HUB.getLastRow()
  let masterRosterRosterLength = MASTER_ROSTER.getLastRow();
  
  return masterRosterRosterLength - emailHubRosterLength;  
}

function getGuardianEmails(studentEmailToLookUp) {
  if(IMPORT_GUARDIAN_EMAILS_FROM == "") {return ["",""];} 
  let importFromThisSpreadsheet = SpreadsheetApp.openById(IMPORT_GUARDIAN_EMAILS_FROM);
  let sheets = importFromThisSpreadsheet.getSheets();
  
  let sheetHeader = sheets[0].getRange(1,1,1,sheets[0].getLastColumn()).getValues();
  
  let studentEmails = sheets[0].getRange(2,getIndexOf(sheetHeader, EMAIL_COLUMN_NAME),sheets[0].getLastRow(),1).getValues().flat(1);
  
  let matchIndex = studentEmails.indexOf(studentEmailToLookUp);
  
  if(matchIndex == -1) return ["MATCH NOT FOUND", "MATCH NOT FOUND"];
  else {
    let guardianEmail1 = sheets[0].getRange(matchIndex+2, getIndexOf(sheetHeader, GUARDIAN_ONE_COLUMN_NAME)).getDisplayValue();
    let guardianEmail2 = sheets[0].getRange(matchIndex+2, getIndexOf(sheetHeader, GUARDIAN_TWO_COLUMN_NAME)).getDisplayValue();
    
    if(guardianEmail1 == "") {guardianEmail1 = "NOT PROVIDED";}
    if(guardianEmail2 == "") {guardianEmail2 = "NOT PROVIDED";}
    
    return [guardianEmail1, guardianEmail2];
  }
  
}


function generateEmailDrafts() {
  let ccParents = copyParentsPrompt();
  if(ccParents == END_FUNCTION) return;
  
  actionCompleteNotifier("","");
  
  var EMAIL_TEMPLATE = DocumentApp.openById(EMAIL_TEMPLATE_ID).getText();
  let header = EMAIL_HUB.getSheetValues(1,1,1,EMAIL_HUB.getLastColumn());
  let draftsAlreadyCreated = GmailApp.getDrafts();
  
  let listOfMissingAssignments = MASTER_ROSTER.getRange(2,7,MASTER_ROSTER.getLastRow()-1,1).getValues().flat(1);
  
  listOfMissingAssignments.forEach((list, index) => {
    //list = list.replace(/(\r\n|\n|\r)/gm, "");
    let splitList = list.split(MISSING_ASSIGNMENT_LIST_DELIMITER);
    let missingExtraWork = splitList[2];
    let missingClasswork = splitList[1];
    let missingHomework = splitList[0];
    
    let missingMandatoryWork = missingHomework.concat(missingClasswork);
    
    let row = index + 2;
    
    if(missingMandatoryWork != "\n " && missingMandatoryWork.length > 2) {
      let replaceWithThese = [];
      let formattedStudentName = makeFirstLetterUpperCase(EMAIL_HUB.getRange(row,getIndexOf(header,FN_COLUMN_NAME)).getDisplayValue());
      replaceWithThese.push(formattedStudentName);
      
      let ratioMissingAssignments = MASTER_ROSTER.getRange(row,5).getDisplayValue();
      let numMissingAssignments = ratioMissingAssignments.split(':');
      replaceWithThese.push(numMissingAssignments[0]);
      
      let formattedClassName = formatCourseName(EMAIL_HUB.getRange(row, getIndexOf(header, CLASS_COLUMN_NAME)).getDisplayValue());
      replaceWithThese.push(formattedClassName);
      
      replaceWithThese.push(missingMandatoryWork); //missing assignment list
      
      let dueDate = EMAIL_HUB.getRange(2, getIndexOf(header, DEADLINE_COLUMN_NAME)).getDisplayValue();
      replaceWithThese.push(dueDate);
     
      let parentsEmails = ","
      if(ccParents == PROCEED) {
        let listOfErrorMessages = ["NOT PROVIDED", "MATCH NOT FOUND"]
        let parentEmail1 = EMAIL_HUB.getRange(row,getIndexOf(header,GUARDIAN_ONE_COLUMN_NAME)).getDisplayValue();
        let parentEmail2 = EMAIL_HUB.getRange(row,getIndexOf(header,GUARDIAN_TWO_COLUMN_NAME)).getDisplayValue();
        
        if(!listOfErrorMessages.includes(parentEmail1) && !listOfErrorMessages.includes(parentEmail2)) parentsEmails = parentEmail1 + "," + parentEmail2;
        else if(!listOfErrorMessages.includes(parentEmail1)) parentsEmails = parentEmail1;
        else if(!listOfErrorMessages.includes(parentEmail2)) parentsEmails = parentEmail2;     
      }
      
      let replacementEmailBody = EMAIL_TEMPLATE;
      let emailSections = EMAIL_TEMPLATE.split(EMAIL_TEMPLATE_BREAK_IDENTIFIER);
      
      [generalEmail, optionalBoost, parentsIncluded, closingSignature] = emailSections;
      
      if(missingExtraWork == "" && ccParents == PROCEED) {
        replacementEmailBody = generalEmail.concat(parentsIncluded, closingSignature);
      }
      else if(missingExtraWork == "" && ccParents == DECLINE) {
        replacementEmailBody = generalEmail.concat(closingSignature);
      }
      else if(missingExtraWork != "" && ccParents == DECLINE) {
        replacementEmailBody = generalEmail.concat(optionalBoost, closingSignature);
        replaceWithThese.push(missingExtraWork); 
      }
      else if(missingExtraWork != "" && ccParents == PROCEED) {
        replacementEmailBody = generalEmail.concat(optionalBoost, parentsIncluded, closingSignature);
        replaceWithThese.push(missingExtraWork); //missing optional assignments
        Logger.log(formattedStudentName, missingExtraWork);
      }
      
      replaceWithThese.forEach(replacement => {
        replacementEmailBody = replacementEmailBody.replace(/{[^{}]+}/i,replacement)
      });
      
      
      let emailSubject = EMAIL_SUBJECT.concat(makeFirstLetterEachWordUpper(formattedClassName));
      
      let currentStudentDraftId = EMAIL_HUB.getRange(row,getIndexOf(header, DRAFT_ID_COLUMN_NAME)).getDisplayValue();
      let updateThisDraft = draftsAlreadyCreated.find(draft => draft.getId() == currentStudentDraftId);
     
      let studentEmailAddress = EMAIL_HUB.getRange(row,getIndexOf(header, EMAIL_COLUMN_NAME)).getDisplayValue();
      if(currentStudentDraftId == "" || currentStudentDraftId == "N/A" || updateThisDraft == undefined) {
        let draft = "";
        if(parentsEmails != ",") draft = GmailApp.createDraft(studentEmailAddress,emailSubject, replacementEmailBody, {cc: parentsEmails});
        else draft = GmailApp.createDraft(studentEmailAddress,emailSubject, replacementEmailBody);
        EMAIL_HUB.getRange(row,getIndexOf(header, DRAFT_ID_COLUMN_NAME)).setValue(draft.getId());
      }
      else {
        let updatedBody = replacementEmailBody;
        if(parentsEmails != ",") updateThisDraft.update(studentEmailAddress, emailSubject, updatedBody, {cc: parentsEmails});
        else updateThisDraft.update(studentEmailAddress, emailSubject, updatedBody);
      }
    }
    else 
    {
      EMAIL_HUB.getRange(row,getIndexOf(header, DRAFT_ID_COLUMN_NAME)).setValue("N/A");
    }
  });

  EMAIL_HUB.getRange(2,getIndexOf(header, DATE_DRAFTED_COLUMN_NAME)).setValue(new Date());
  actionCompleteNotifier("Finished!","");
}

function copyParentsPrompt(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Generate Emails', 'Do you want to copy parents on email?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) return PROCEED;
  else if(response == ui.Button.NO) return DECLINE;
  else return END_FUNCTION;
}

function makeFirstLetterUpperCase(string) {
  return string.charAt(0).toUpperCase() + string.slice(1).toLowerCase();
}

function makeFirstLetterEachWordUpper(string) {
  let words = string.split(' ');
  let upperCaseWords = "";
  words.forEach((word) => upperCaseWords = upperCaseWords.concat(" ", makeFirstLetterUpperCase(word)));
  return upperCaseWords;
  
}

function formatCourseName(courseName) {
  COURSE_DICTIONARY.forEach(keyValuePair => {
    if(courseName.toLowerCase().includes(keyValuePair.key)) {
      courseName = keyValuePair.value;
      return;
    }
  });
    
  return courseName;
}

function formatTopicRow(sheet) {
  let topicRow = sheet.getRange(1,4,1,sheet.getLastColumn()).getDisplayValues().flat(1);
  let arrayOfIndexes = [];
  for(i = 0; i < topicRow.length; i++) {
    if(topicRow[i] != "") arrayOfIndexes.push(i+4);
  }
 
  for(j = 0; j < arrayOfIndexes.length; j++) {
    if(j == arrayOfIndexes.length-1) {
      Logger.log(arrayOfIndexes[j], topicRow.length)
      let range = sheet.getRange(1,arrayOfIndexes[j],1,topicRow.length-arrayOfIndexes[j]+1)
      range.merge();
      range.setHorizontalAlignment("center");
      range.setVerticalAlignment("middle");
      range.setWrap(true);
    }
    else {
      Logger.log(arrayOfIndexes[j], arrayOfIndexes[j+1]-1);
      let range = sheet.getRange(1,arrayOfIndexes[j],1,arrayOfIndexes[j+1]-arrayOfIndexes[j])
      range.merge();
      range.setHorizontalAlignment("center");
      range.setVerticalAlignment("middle");
      range.setWrap(true);    
    }
    
  }

}

function formatSheets() {
  [a, b, ...CLASS_LIST] = GRADE_REPORT.getSheets();
  CLASS_LIST.forEach((sheet) => { 
      formatTopicRow(sheet);
      let sortByStudentLastNameRange = sheet.getRange(5,1,sheet.getLastRow(),sheet.getLastColumn());
      sortByStudentLastNameRange.sort(2);
      
      sheet.setFrozenColumns(2);
      
      let topicRange = sheet.getRange(1,4,1,sheet.getLastColumn());
      let topicStyle = SpreadsheetApp.newTextStyle()
        .setFontSize(16)
        .setBold(true)
        .build();
      topicRange.setTextStyle(topicStyle);
      topicRange.setWrap(true);

      let assignmentRange = sheet.getRange(2,4,1,sheet.getLastColumn());
      assignmentRange.setWrap(true);
      assignmentRange.setHorizontalAlignment("center");
      assignmentRange.setVerticalAlignment("middle");
      
      sheet.autoResizeColumn(3);
  });

  createConditionalFormatting();
}

function formatEmailHub(){
  EMAIL_HUB.clearFormats();
  
  EMAIL_HUB.setFrozenRows(1);
  EMAIL_HUB.setFrozenColumns(3);
  
  EMAIL_HUB.autoResizeColumn(4);
  
  sortAndMergeSheet(EMAIL_HUB, true, false);
}

function formatMasterRoster(){
  MASTER_ROSTER.clearFormats();
  
  MASTER_ROSTER.setFrozenRows(1);
  MASTER_ROSTER.setFrozenColumns(3);
  
  MASTER_ROSTER.autoResizeColumn(4);
  
  sortAndMergeSheet(MASTER_ROSTER, true, true);
  
  let missingAssignmentListRange = [];
  let missingAssignmentRange = [];
  for(i = 0; i < NUM_TOPICS_TO_VIEW; i++) {
    MASTER_ROSTER.setColumnWidth(7+3*i,188);
    let missingAssignmentListRangeA = MASTER_ROSTER.getRange(2,7+2*i,MASTER_ROSTER.getLastRow(),1);
    missingAssignmentListRange.push(missingAssignmentListRangeA);

    let missingAssignmentRangeA = MASTER_ROSTER.getRange(2,5+3*i, MASTER_ROSTER.getLastRow(), 2);
    missingAssignmentRange.push(missingAssignmentRangeA);
  }

  missingAssignmentListRange.forEach((range) => range.setWrap(true)); 
  missingAssignmentRange.forEach((range) => range.setHorizontalAlignment("left"));
  
  let noMissingAssignmentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith(0)
    .setBackground("#d9ead3")
    .setRanges(missingAssignmentRange)
    .build();
    
  let acceptableNummissingHomeworkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEFT(E2,SEARCH(":",E2)-1) < RIGHT(E2,SEARCH(":",E2)-1)')
    .setBackground("#fce5cd")
    .setRanges(missingAssignmentRange)
    .build();
 
   let excessiveNummissingHomeworkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEFT(E2,SEARCH(":",E2)-1) >= RIGHT(E2,SEARCH(":",E2)-1)')
    .setBackground("#f4cccc")
    .setRanges(missingAssignmentRange)
    .build();
    
   let rules = MASTER_ROSTER.getConditionalFormatRules()
   rules.push(noMissingAssignmentRule, acceptableNummissingHomeworkRule, excessiveNummissingHomeworkRule);
   MASTER_ROSTER.setConditionalFormatRules(rules)
}

function createConditionalFormatting() {
  CLASS_LIST.forEach((sheet) => {
    let topicRange = sheet.getRange(1,4,sheet.getLastRow(),sheet.getLastColumn())
    let topicRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=NOT(ISBLANK(D$1))')
    .setBackground("#efefef")
    .setRanges([topicRange])
    .build()
    
    let assignmentNameRange = sheet.getRange(2,4,1,sheet.getLastColumn());
    let assignmentNameRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=NOT(ISBLANK(D$2))')
    .setBackground("#9fc5e8")
    .setRanges([assignmentNameRange])
    .build()
    
    let assignmentDueDateRange = sheet.getRange(3,4,1,sheet.getLastColumn());
    let assignmentDueDateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=NOT(ISBLANK(D$3))')
    .setBackground("#cfe2f3")
    .setRanges([assignmentDueDateRange])
    .build()
    
    let submissionRange = sheet.getRange(4,5,sheet.getLastRow(),sheet.getLastColumn())
    let notSubmittedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("MISSING")
    .setBackground("#cc0000")
    .setRanges([submissionRange])
    .build()
    
    let submittedOnTimeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('SUBMITTED')
    .setBackground("#b8eba6")
    .setRanges([submissionRange])
    .build()
    
    let submittedLateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ARRAYFORMULA(SPLIT(E4, "-") >= E$3)')
    .setBackground("#ffb870")
    .setRanges([submissionRange])
    .build()
    
    let gradedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("GRADED")
    .setBackground("#d9ead3")
    .setRanges([submissionRange])
    .build()
    
    let notYetDueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("NOT YET DUE")
    .setBackground("#f9d0a6")
    .setRanges([submissionRange])
    .build()
    
    let extraCreditRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("EXTRA CREDIT")
    .setBackground("#cfe2f3")
    .setRanges([submissionRange])
    .build()
    
    let optionalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("OPTIONAL")
    .setBackground("#cfe2f3")
    .setRanges([submissionRange])
    .build()
    
    let classWorkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("NOT YET SUB")
    .setBackground("#f4cccc")
    .setRanges([submissionRange])
    .build()

      
    let rules = sheet.getConditionalFormatRules()
    rules.push(topicRule, assignmentNameRule, assignmentDueDateRule, gradedRule, notSubmittedRule, submittedOnTimeRule, notYetDueRule, extraCreditRule, classWorkRule, submittedLateRule, optionalRule)
    sheet.setConditionalFormatRules(rules)
  });
}

function testSortandMerge() {
  sortAndMergeSheet(MASTER_ROSTER, false, false);
  


}

function sortAndMergeSheet(sheetToFormat, wantToSort, wantToMerge){
  let classColumn = sheetToFormat.getRange(1,1,sheetToFormat.getLastRow(),1);
  let classValues = classColumn.getValues().flat(1);
  let classesToBeMerged = classValues.filter((cell) => cell != ""); 
  
  CLASS_LIST.forEach((sheet) => {
    let sheetName = sheet.getName().split(" (");
    let filteredClassesBySheetName = classesToBeMerged.filter((cell) => cell == sheetName[0]);
    let startPosition = classesToBeMerged.indexOf(sheetName[0])+1;
    
    if(filteredClassesBySheetName.length > 0) {
      let numOfColumns = 0;
      if(sheetToFormat == EMAIL_HUB) numOfColumns = 7;
      if(sheetToFormat == MASTER_ROSTER) numOfColumns = MASTER_ROSTER.getLastColumn();
      let sortRange = sheetToFormat.getRange(startPosition, 1, filteredClassesBySheetName.length, numOfColumns);
      
      if(wantToSort) sortRange.sort(3);
      if(wantToMerge) mergeAndCenterCells(MASTER_ROSTER, startPosition, 1, filteredClassesBySheetName.length, 1); 
    }
  })
}




