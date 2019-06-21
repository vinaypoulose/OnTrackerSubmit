function onSubmit(e) {

  var emailBody = "";

  try {

    var response = e.response;
    var responseId = response.getId();
    var timeStamp = response.getTimestamp();
    var submitter = response.getRespondentEmail();
    var userSubmittedTrackerId = response.getItemResponses()[0].getResponse().toString();

    //    var responseId = 12345678;
    //    var timeStamp = new Date();
    //    var submitter = "vinay.poulose@akanksha.org";
    //    var userSubmittedTrackerId = "1OJvg8KbQqVat3qyHoOytEfdCTqc78Q53";

    var urlUserFile = "https://drive.google.com/open?id=" + userSubmittedTrackerId;

    //Utilities.sleep(5000);

    var googleSheetWithUserSubmittedContentId = CreateGoogleSpreadSheet(userSubmittedTrackerId);

    emailBody += "googleSheetWithUserSubmittedContentId: " + googleSheetWithUserSubmittedContentId + "\n";

    emailBody += onTrackerSubmit(timeStamp, responseId, submitter, googleSheetWithUserSubmittedContentId, urlUserFile);

    MailApp.sendEmail("vinay.poulose@akanksha.org", "vinay.poulose@akanksha.org", "D3f: Successful data upload", emailBody);

  } catch (exception) {

    MailApp.sendEmail("vinay.poulose@akanksha.org", "vinay.poulose@akanksha.org", "D3f: Failed data upload",
      "There was an error while trying to add your data to D3f!\n" + exception.toString());
  }
}

function onTrackerSubmit(timeStamp, responseId, submitter, idGoogleSheetUserFileConverted, urlUserFile) {

  const urlSchoolDashboard = "https://docs.google.com/spreadsheets/d/1-natPhgAxf9Q8fwbg8wbSV_OQT6pSvr7RyhvgApFUEU/edit#gid=403349783";
  const urlTemplateDashboard = "https://docs.google.com/spreadsheets/d/19RQUtQ6G3LrFaHUcF77a4_7AI0xoLlk4uiTqnMATd3w/edit#gid=0";
  const urlTemplateGoogleSheetScoreBased = "https://docs.google.com/spreadsheets/d/1LEDErVIqqHJxXT_qFziJDTFYoYUYTD0XjUJ_eQTIfzQ/edit#gid=0";
  const rowOffsetMarksObtained = 42;
  const colOffsetMarksObtained = 41;
  const rowTrackerMaxMarks = 35;
  const rowOffsetSkills = 4;

  var urlDashboard = "https://docs.google.com/spreadsheets/d/1zZ0QcsZo8KijckgeCwGZc2Kwu2qdojTJxB5ZdGWRmtg/edit#gid=835067691";


  //  timeStamp = new Date();
  //  responseId = 123456789;
  //  submitter = "d3f.demo@akanksha.org";
  //  idGoogleSheetUserFileConverted = "1kM0oMsuCN4cu7qwcyKYhphzDAa8omUiHv7h_HgSv3tg";
  //  urlUserFile = "drive.google.com";


  var i, j;
  var isSchoolValid = false;
  var isStandardValid = false;
  var isDivisionValid = false;
  var isSubjectValid = false;
  var isDateValid = false;
  var trackerValid = true;


  var emailBody = [];

  //  var googleSheetWithUserSubmittedContentId = CreateGoogleSpreadSheet(userSubmittedTrackerId);
  var idGoogleSheetWithUserSubmittedContent = "14vpU4VaOfkHoF36vr3vHoxlf4MLy4hVTPL7EmZGSM_k";

  if (idGoogleSheetUserFileConverted == null || idGoogleSheetUserFileConverted == "") {

    emailBody.push("There was an error while trying to add the assessment data you submitted to D3:f\nCould not convert Google Sheet to Excel file");
    return emailBody.join("\n");
  }

  var googleSheetWithUserSubmittedContent = SpreadsheetApp.openById(idGoogleSheetUserFileConverted);

  if (googleSheetWithUserSubmittedContent == null) {

    emailBody.push("There was an error while trying to add the assessment data you submitted to D3:f\nCould not open converted Google Sheet");
    return emailBody.join("\n");
  }

  var sheetScores = googleSheetWithUserSubmittedContent.getSheetByName("scores");

  if (sheetScores == null) {

    emailBody.push("There was an error while trying to add the assessment data you submitted to D3:f\nCould not open converted Google Sheet");
    return emailBody.join("\n");
  }

  var studentDb = googleSheetWithUserSubmittedContent.getRangeByName("tableStudentDb").getValues();
  var listSubjects = googleSheetWithUserSubmittedContent.getRangeByName("listSubjects").getValues();

  var school = googleSheetWithUserSubmittedContent.getRangeByName("school").getValue();
  var standard = googleSheetWithUserSubmittedContent.getRangeByName("standard").getValue();
  var division = googleSheetWithUserSubmittedContent.getRangeByName("division").getValue();
  var subject = googleSheetWithUserSubmittedContent.getRangeByName("subject").getValue();
  var date = googleSheetWithUserSubmittedContent.getRangeByName("date").getValue();
  var cceComponent = googleSheetWithUserSubmittedContent.getRangeByName("cceComponent").getValue();
  var tag = googleSheetWithUserSubmittedContent.getRangeByName("tag").getValue();
  var isRubricBasedAssessment = googleSheetWithUserSubmittedContent.getRangeByName("isRubricBasedAssessment").getValue();
  var isTermEndAssessment = googleSheetWithUserSubmittedContent.getRangeByName("isTermEndAssessment").getValue();
  var includeInReportCard = googleSheetWithUserSubmittedContent.getRangeByName("includeInReportCard").getValue();

  //var isRubricBasedAssessment = googleSheetWithUserSubmittedContent.getRangeByName("tableUserEnteredValues").getValue();

  var userEditingRange = googleSheetWithUserSubmittedContent.getRangeByName("tableUserEnteredValues");
  var userEnteredValues = userEditingRange.offset(0, 0, userEditingRange.getLastRow(), userEditingRange.getLastColumn()).getValues();

  var tableStudentDetails = [];

  for (i = 0; i < studentDb.length; i++) {

    var schoolFromStudentDb = studentDb[i][5];
    var standardFromStudentDb = studentDb[i][9];
    var divisionFromStudentDb = studentDb[i][7];

    if (schoolFromStudentDb != null && schoolFromStudentDb.toString().toLowerCase() == school.toString().toLowerCase()) {

      isSchoolValid = true;

      if (standardFromStudentDb != null && standardFromStudentDb.toString().toLowerCase() == standard.toString().toLowerCase()) {

        isStandardValid = true;

        if (divisionFromStudentDb != null && divisionFromStudentDb.toString().toLowerCase() == division.toString().toLowerCase()) {

          isDivisionValid = true;

          var studentId = studentDb[i][0];
          var firstName = studentDb[i][1];
          var lastName = studentDb[i][2];
          var gender = studentDb[i][3];
          var dateOfBirth = studentDb[i][4];

          tableStudentDetails.push([studentId, firstName + " " + lastName, null, null, null, null, dateOfBirth, gender, school, standard, division]);
        }
      }
    }
  }


  for (i = 0; i < listSubjects.length; i++) {

    if (listSubjects[i] != null && listSubjects[i].length > 0 && listSubjects[i][0].toString().toLowerCase() == subject.toString().toLowerCase()) {

      isSubjectValid = true;
      break;
    }
  }

  var schoolDashboard = SpreadsheetApp.openByUrl(urlSchoolDashboard);
  var lockingPeriodInDays  = schoolDashboard.getRangeByName("lockingPeriodInDays").getValue();

  isDateValid = isValidDate(date, lockingPeriodInDays, emailBody);

  if (cceComponent == null || cceComponent.toString() == "") {

    cceComponent = "Not Defined";
  }

  if (tag == null || tag.toString() == "") {

    tag = "Not Defined";
  }

  if (isRubricBasedAssessment != null && isRubricBasedAssessment == true) {

    isRubricBasedAssessment = "Rubric";

  } else {

    isRubricBasedAssessment = "Score";

  }

  if (isTermEndAssessment != null && isTermEndAssessment.toString().toLowerCase() == "yes") {

    isTermEndAssessment = "Summative";

  } else {

    isTermEndAssessment = "Formative";
  }

  if (includeInReportCard != null && includeInReportCard.toString().toLowerCase() == "no") {

    includeInReportCard = false;

  } else {

    includeInReportCard = true;
  }

  

  trackerValid = trackerValid && getErrorForMandatoryField(school, isSchoolValid, "school", emailBody);
  trackerValid = trackerValid && getErrorForMandatoryField(subject, isSubjectValid, "subject", emailBody);
  trackerValid = trackerValid && getErrorForMandatoryField(standard, isStandardValid, "standard", emailBody);
  trackerValid = trackerValid && getErrorForMandatoryField(division, isDivisionValid, "division", emailBody);
  trackerValid = trackerValid && isDateValid;

  if (trackerValid == false) {

    return emailBody.join("\n");
  }

  
  var assessmentId = makeKey([school, standard, division, subject, "" + Utilities.formatDate(date, "IST", "YYYYMMdd")]);
  var arrLog = [];
  var arrRowsToInsertInAssessmentDb = [];
  var arrAssessmentDetails = [assessmentId, subject, date, cceComponent, tag, isRubricBasedAssessment, isTermEndAssessment, includeInReportCard];


  var permissionsDb = schoolDashboard.getRangeByName("permissionData").getValues();
  var rowPermissionForClass;
  var editorForSubjectForClass;

  for (i = 0; i < permissionsDb.length; i++) {

    if (permissionsDb[i][0] == standard && permissionsDb[i][1] == division) {

      rowPermissionForClass = permissionsDb[i];
      break;
    }
  }

  if (rowPermissionForClass != null) {

    for (j = 2; j < rowPermissionForClass.length && j < permissionsDb[0].length; j++) {

      if (permissionsDb[0][j] == subject) {

        editorForSubjectForClass = rowPermissionForClass[j];

      }
    }
  } else {

    emailBody.push("Cannot retrieve the tracker submission permissions for " + school + " " + standard + division + " " + subject);
    return emailBody.join("\n");
  }

  if (submitter.toString().toLowerCase() != editorForSubjectForClass.toString().toLowerCase()) {

    emailBody.push("You (" + submitter + ")have not been assigned permission to submit trackers for " + school + " " + standard + division + " " + subject);
    return emailBody.join("\n");
  }

  //  trackerValid = trackerValid && validateScoreBasedTracker(tableStudentDetails, userEnteredValues, rowOffsetMarksObtained,
  //                                                          colOffsetMarksObtained, rowOffsetSkills, rowTrackerMaxMarks,
  //                                                          arrRowsToInsertInAssessmentDb, [timeStamp, responseId],
  //                                                          arrAssessmentDetails, arrLog);

  if (isRubricBasedAssessment == "Score") {

    trackerValid = trackerValid && validateScoreBasedTracker(tableStudentDetails, userEnteredValues, rowOffsetMarksObtained,
      colOffsetMarksObtained, rowOffsetSkills, rowTrackerMaxMarks,
      arrRowsToInsertInAssessmentDb, [timeStamp, responseId],
      arrAssessmentDetails, emailBody);

  } else if (isRubricBasedAssessment == "Rubric") {

    var listRubric = googleSheetWithUserSubmittedContent.getRangeByName("listRubric").getValues();

    for (i = listRubric.length; i > 0; i--) {

      if (listRubric[i - 1][0].toString() == null || listRubric[i - 1][0].toString() == "") {

        listRubric.pop();
      }
    }

    trackerValid = trackerValid && validateRubricBasedTracker(tableStudentDetails, listRubric, userEnteredValues, rowOffsetMarksObtained,
      colOffsetMarksObtained, rowOffsetSkills, rowTrackerMaxMarks,
      arrRowsToInsertInAssessmentDb, [timeStamp, responseId],
      arrAssessmentDetails, emailBody);

  } else {

    emailBody.push("Cannot determine if this is a rubric or score based tracker");
    return emailBody.join("\n");

  }

  if (trackerValid == true) {

    var dashboardExists = false;
    var urlDashboard;
    var googleSheetDashboard;

    var acadYear = "";
    var year = date.getYear();
    var month = date.getMonth();

    if (month >= 4) {

      acadYear = year + "_" + (year + 1).toString().substr(2, 2);

    } else {

      acadYear = (year - 1) + "_" + year.toString().substr(2, 2);
    }

    var rangeDashboardDb = schoolDashboard.getRangeByName("dashBoardData");
    var dashboardDb = rangeDashboardDb.getValues();

    for (i = 0; i < dashboardDb.length; i++) {

      if (dashboardDb[i][0] == standard && dashboardDb[i][1] == division && dashboardDb[i][2] == subject) {

        urlDashboard = dashboardDb[i][3];
        if (urlDashboard != null && urlDashboard.toString() != "") {

          dashboardExists = true;
        }
        break;
      }
    }

    if (dashboardExists == false) {

      googleSheetDashboard = SpreadsheetApp.openByUrl(urlTemplateDashboard).copy((acadYear + "_" + school + "_" + standard + "_" + division + "_" + subject).toLowerCase());
      googleSheetDashboard.getRangeByName("assessmentDetailsStartCell").offset(0, 0, 4, 1).setValues([[school], [standard], [division], [subject]]);
      DriveApp.getFileById(googleSheetDashboard.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      urlDashboard = googleSheetDashboard.getUrl();
      rangeDashboardDb.getSheet().appendRow([standard, division, subject, urlDashboard]);

    } else {

      googleSheetDashboard = SpreadsheetApp.openByUrl(urlDashboard);
    }

    if (googleSheetDashboard == null) {

      emailBody.push("There was an error trying to open the " + subject + " dashboard for " + school + " " + standard + " " + division);
      emailBody.push("URL: " + urlDashboard);

      return emailBody.join("\n");;
    }

    var db = googleSheetDashboard.getRangeByName("assessmentData");
    var sheetDb = db.getSheet();
    var dbRows = db.getValues();

    if (dbRows == null) {

      emailBody.push("There was an error trying to get values from the range named 'assessmentData' from the " + subject + " dashboard for " + school + " " + standard + " " + division);

      return emailBody.join("\n");;
    }

    var trackerDb = googleSheetDashboard.getRangeByName("trackerData");
    var sheetTrackerDb = trackerDb.getSheet();
    var trackerDbRows = trackerDb.getValues();

    if (dbRows == null) {

      emailBody.push("There was an error trying to get values from the range named 'trackerData' from the " + subject + " dashboard for " + school + " " + standard + " " + division);

      return emailBody.join("\n");;
    }

    var trackerExists = false, trackerCount = 0;
    var urlGoogleSheetTracker, googleSheetTracker;

    for (i = 0; i < trackerDbRows.length; i++) {

      var row = trackerDbRows[i];

      if (trackerDbRows[i][13].toString().toLowerCase() == "active") {

        trackerCount++;

        if (trackerDbRows[i][2] == assessmentId) {
          trackerExists = true;
          urlGoogleSheetTracker = trackerDbRows[i][3];

          if (row[10] != isRubricBasedAssessment) {

            emailBody.push("You have previously submitted a " + row[10] + " based tracker for this assessment. Cannot change it to a " + isRubricBasedAssessment + " based tracker.");
            return emailBody.join("\n");
          }

          row[8] = cceComponent;
          row[9] = tag;
          row[10] = isRubricBasedAssessment;
          row[11] = isTermEndAssessment;
          row[12] = includeInReportCard;

          sheetTrackerDb.getRange(i + 1, 1, 1, 14).setValues([row]);

          break;
        }
      }
    }

    var googleSheetTracker;

    if (trackerExists == true) {

      googleSheetTracker = SpreadsheetApp.openByUrl(urlGoogleSheetTracker);

    } else {

      var templateGoogleSheetScoreBased = SpreadsheetApp.openByUrl(urlTemplateGoogleSheetScoreBased);
      googleSheetTracker = templateGoogleSheetScoreBased.copy((acadYear + "_" + school + "_" + standard + "_" + division + "_" + subject + "_" + Utilities.formatDate(date, "IST", "MMM_d")).toLowerCase());
      DriveApp.getFileById(googleSheetTracker.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      urlGoogleSheetTracker = googleSheetTracker.getUrl();
    }

    var assessmentDetails = [[school], [standard], [division], [subject], [date], [cceComponent], [tag], [isTermEndAssessment ? "YES" : "NO"], [includeInReportCard ? "YES" : "NO"]];
    googleSheetTracker.getRangeByName("assessmentDetailsStartCell").offset(0, 0, assessmentDetails.length, 1).setValues(assessmentDetails);

    googleSheetTracker.getRangeByName("studentDetailsStartCell").clearContent();
    googleSheetTracker.getRangeByName("studentDetailsStartCell").offset(0, 0, tableStudentDetails.length, tableStudentDetails[0].length).setValues(tableStudentDetails);

    googleSheetTracker.getRangeByName("tableUserEnteredValues").clearContent();
    googleSheetTracker.getRangeByName("tableUserEnteredValues").offset(0, 0, userEnteredValues.length, userEnteredValues[0].length).setValues(userEnteredValues);

    googleSheetTracker.getRangeByName("auditLogDb").getSheet().appendRow([timeStamp, responseId, submitter, urlUserFile]);

    if (trackerExists == false) {

      sheetTrackerDb.appendRow([timeStamp, responseId, assessmentId, urlGoogleSheetTracker, subject, date, standard, division, cceComponent, tag, isRubricBasedAssessment, isTermEndAssessment, includeInReportCard, "active"]);
      schoolDashboard.getRangeByName("trackerData").getSheet().appendRow([timeStamp, responseId, assessmentId, urlGoogleSheetTracker, subject, date, standard, division, cceComponent, tag, isRubricBasedAssessment, isTermEndAssessment, includeInReportCard, "active"]);

    } else {

      var rowCount = 0, startRow;
      var rowsToInactivate = [];

      for (i = 0; i < dbRows.length; i++) {

        if (dbRows[i][9] == assessmentId && dbRows[i][22] == "active") {

          if (startRow == null) startRow = i + 1;
          rowCount++;

          rowsToInactivate.push(["inactive"]);
        }
      }

      if (startRow != null && startRow > 0 && rowsToInactivate.length > 0) sheetDb.getRange(startRow, 23, rowsToInactivate.length, 1).setValues(rowsToInactivate);

      var trackerDbInSchoolDashboard = schoolDashboard.getRangeByName("trackerData").getValues();

      if (trackerDbInSchoolDashboard != null) {

        for (i = 0; i < trackerDbInSchoolDashboard.length; i++) {

          var currentRow = trackerDbInSchoolDashboard[i];

          if (currentRow[2] == assessmentId && currentRow[13].toString().toLowerCase() == "active") {

            currentRow[8] = cceComponent;
            currentRow[9] = tag;
            currentRow[10] = isRubricBasedAssessment;
            currentRow[11] = isTermEndAssessment;
            currentRow[12] = includeInReportCard;

            schoolDashboard.getRangeByName("trackerData").getSheet().getRange(i + 1, 1, 1, 14).setValues([currentRow]);

            break;
          }
        }
      }
    }

    var rowNumLastInDb = sheetDb.getLastRow();
    var rowNumInsertAfter = rowNumLastInDb;
    var numRowsToInsert = arrRowsToInsertInAssessmentDb.length;

    if (rowNumLastInDb == 0) {

      rowNumInsertAfter++;
      numRowsToInsert--;

    }

    sheetDb.insertRowsAfter(rowNumInsertAfter, numRowsToInsert);
    sheetDb.getRange(rowNumLastInDb + 1, 1, arrRowsToInsertInAssessmentDb.length, arrRowsToInsertInAssessmentDb[0].length).setValues(arrRowsToInsertInAssessmentDb);

    SpreadsheetApp.flush();

    emailBody.push("Assessment data submitted by you has been successfully uploaded. You can view the analysis here:");
    emailBody.push(urlGoogleSheetTracker);

  }

  return emailBody.join("\n");
}


function validateScoreBasedTracker(tableStudentDetails, tableMarksObtained, rowOffsetMarksObtained,
  colOffsetMarksObtained, rowOffsetSkills, rowTrackerMaxMarks,
  arrRowsToInsertInAssessmentDb, arrSubmissionDetails, arrAssessmentDetails, arrLog) {

  var i, isTrackerValid = true;
  var arrMaxMarks = null;
  var arrIsMaxMarksValid = [];
  var arrIsMaxMarksEmpty = [];
  var arrSkills = [];
  var rowNumHeadersInMarksObtained = rowOffsetMarksObtained - rowTrackerMaxMarks;

  var studentId;
  var isStudentIdAvailable;
  var attendance;
  var isStudentPresent;
  var arrMarksObtained;

  var dbRow;
  var maxMarks;
  var isMaxMarksValid;
  var isMaxMarksEmpty;
  var skill;
  var marksObtained;
  var percentage;


  if (tableMarksObtained != null && tableMarksObtained.length > 0) {

    if (tableMarksObtained[0] != null) {

      arrMaxMarks = tableMarksObtained[0].slice(1);
    }

    isTrackerValid = validateMaxMarks(arrMaxMarks, arrIsMaxMarksValid, arrIsMaxMarksEmpty, colOffsetMarksObtained, rowTrackerMaxMarks, arrLog);

    if (rowOffsetSkills < tableMarksObtained.length && tableMarksObtained[rowOffsetSkills] != null) {

      var temp = tableMarksObtained[rowOffsetSkills].slice(1);

      if (temp != null) {

        for (j = 0; j < temp.length; j++) {

          if (temp[j] == null || temp[j].toString() == "") {

            arrSkills.push("not defined");

          } else {

            arrSkills.push(trimAndToLowerCaseAndRemoveExtraSpaces(temp[j]));
          }
        }
      }
    }

    for (i = 0; i < tableMarksObtained.length; i++) {

      studentId = null;
      isStudentIdAvailable = false;
      attendance = null
      isStudentPresent = false;
      arrMarksObtained = null;

      if (tableStudentDetails != null && i < tableStudentDetails.length) {

        studentId = tableStudentDetails[i][0];
        if (studentId != null && studentId.toString() != "") {

          isStudentIdAvailable = true;
        }
      }

      if (tableMarksObtained[i + rowNumHeadersInMarksObtained] != null &&
        tableMarksObtained[i + rowNumHeadersInMarksObtained].length > 0) {

        attendance = tableMarksObtained[i + rowNumHeadersInMarksObtained][0];
        if (attendance != null && attendance.toString().toLowerCase() == "p") isStudentPresent = true;

        arrMarksObtained = tableMarksObtained[i + rowNumHeadersInMarksObtained].slice(1);
      }

      if (arrMarksObtained != null) {

        for (j = 0; j < arrMarksObtained.length; j++) {

          dbRow = arrSubmissionDetails.slice();
          maxMarks = null;
          isMaxMarksValid = false;
          isMaxMarksEmpty = false;
          skill = "not defined";
          percentage = null;

          if (arrMaxMarks != null && j < arrMaxMarks.length) {

            maxMarks = arrMaxMarks[j];
          }

          if (j < arrIsMaxMarksValid.length) {

            isMaxMarksValid = arrIsMaxMarksValid[j];
          }

          if (j < arrIsMaxMarksEmpty.length) {

            isMaxMarksEmpty = arrIsMaxMarksEmpty[j];
          }

          marksObtained = arrMarksObtained[j];

          if (arrSkills != null && j < arrSkills.length) {

            skill = arrSkills[j];
          }

          isTrackerValid = isTrackerValid && validateMarksObtained(maxMarks, isMaxMarksValid, isMaxMarksEmpty, marksObtained, isStudentIdAvailable,
            isStudentPresent, i + rowOffsetMarksObtained, colOffsetMarksObtained, arrLog);


          if (isTrackerValid == true && isStudentIdAvailable == true && isMaxMarksValid == true) {

            if (isStudentPresent == true) {

              percentage = marksObtained / maxMarks;
            }

            dbRow = dbRow.concat([studentId,
              tableStudentDetails[i][1],
              tableStudentDetails[i][6],
              tableStudentDetails[i][7],
              tableStudentDetails[i][8],
              tableStudentDetails[i][9],
              tableStudentDetails[i][10]],
              arrAssessmentDetails,
              [skill,
                attendance,
                maxMarks,
                marksObtained,
                percentage,
                "active"]);


            arrRowsToInsertInAssessmentDb.push(dbRow);
          }
        }

      } else if (isStudentIdAvailable == true) {

        arrLog.push("No marks entered for " + tableStudentDetails[i][1] + " (" + tableStudentDetails[studentId] + ")");
        isTrackerValid == false;
      }
    }
  }

  return isTrackerValid;
}

function validateMaxMarks(arrMaxMarks, arrIsMaxMarksValid, arrIsMaxMarksEmpty, colOffsetMarksObtained, rowTrackerMaxMarks, arrLog) {

  var i, maxMarks, typeMaxMarks, cellAddress, isTrackerValid = true;

  if (arrMaxMarks != null) {

    for (i = 0; i < arrMaxMarks.length; i++) {

      cellAddress = columnToLetter(i + colOffsetMarksObtained) + rowTrackerMaxMarks;

      maxMarks = arrMaxMarks[i];
      typeMaxMarks = typeof maxMarks;

      arrIsMaxMarksValid.push(false);
      arrIsMaxMarksEmpty.push(false);

      if (typeMaxMarks == (typeof 0)) {

        if (maxMarks == 0) {

          arrLog.push("Max Marks entered in " + cellAddress + " is zero");
          isTrackerValid = isTrackerValid && false;

        } else if (maxMarks < 0) {

          arrLog.push("Max Marks entered in " + cellAddress + " is negative");
          isTrackerValid = isTrackerValid && false;

        } else {

          arrIsMaxMarksValid[i] = true;
        }

      } else if (maxMarks == null || maxMarks.toString() == "") {

        arrIsMaxMarksEmpty[i] = true;

      } else {

        arrLog.push("Max Marks entered in " + cellAddress + " is not a valid number");
        isTrackerValid = isTrackerValid && false;
      }
    }
  }

  return isTrackerValid;
}

function validateMarksObtained(maxMarks, isMaxMarksValid, isMaxMarksEmpty, marksObtained, isStudentIdAvailable,
  isStudentPresent, rowNum, colNum, arrLog) {

  var isTrackerValid = true;
  var cellAddress = columnToLetter(colNum) + rowNum;
  var typeMarksObtained = typeof marksObtained;

  if (isStudentIdAvailable == true && isStudentPresent == true) {

    if (isMaxMarksEmpty == false) {

      if (typeMarksObtained == (typeof 0)) {

        if (marksObtained < 0) {

          arrLog.push("Marks entered in " + cellAddress + "is negative");
          isTrackerValid = isTrackerValid && false;

        } else if (isMaxMarksValid == true && marksObtained > maxMarks) {

          arrLog.push("Marks entered in " + cellAddress + "is greater than max marks");
          isTrackerValid = isTrackerValid && false;
        }

      } else if (marksObtained == null || marksObtained.toString() == "") {

        arrLog.push(cellAddress + " is empty");
        isTrackerValid = isTrackerValid && false;

      } else {

        arrLog.push("Marks entered in " + cellAddress + " is not a valid number");
        isTrackerValid = isTrackerValid && false;
      }

    } else if (marksObtained != null && marksObtained.toString() != "") {

      arrLog.push("There is either an extra entry in " + cellAddress + " or the Max Marks has not been entered in column " + columnToLetter(colNum));
      isTrackerValid = isTrackerValid && false;
    }

  } else {

    if (marksObtained != null && marksObtained.toString() != "") {

      arrLog.push("There is an extra entry in " + cellAddress + ". It should be empty");
      isTrackerValid = isTrackerValid && false;
    }
  }

  return isTrackerValid;
}

function getErrorForMandatoryField(fieldValue, isFieldValid, fieldLabel, arrLog) {

  if (isFieldValid == false) {

    if (fieldValue == null || fieldValue.toString() == "") {

      arrLog.push("The " + fieldLabel + " is empty");

    } else {

      arrLog.push("The " + fieldLabel + " entered is invalid");
    }

    return false;
  }

  return true;
}

function validateRubricBasedTracker(tableStudentDetails, listRubric, tableMarksObtained, rowOffsetMarksObtained,
  colOffsetMarksObtained, rowOffsetSkills, rowTrackerMaxMarks,
  arrRowsToInsertInAssessmentDb, arrSubmissionDetails, arrAssessmentDetails, arrLog) {

  var i, j, isTrackerValid = true;
  var arrMaxMarks = null;
  var arrIsMaxMarksValid = [];
  var arrIsMaxMarksEmpty = [];
  var arrSkills = [];
  var rowNumHeadersInMarksObtained = rowOffsetMarksObtained - rowTrackerMaxMarks;

  var studentId;
  var isStudentIdAvailable;
  var attendance;
  var isStudentPresent;
  var arrMarksObtained;

  var dbRow;
  var maxMarks;
  var isMaxMarksValid;
  var isMaxMarksEmpty;
  var skill;
  var marksObtained;
  var percentage;


  if (tableMarksObtained != null && tableMarksObtained.length > 0) {

    if (tableMarksObtained[0] != null) {

      arrMaxMarks = tableMarksObtained[0].slice(1);
    }

    isTrackerValid = validateRubricMaxMarks(listRubric.length, arrMaxMarks, arrIsMaxMarksValid, arrIsMaxMarksEmpty, colOffsetMarksObtained, rowTrackerMaxMarks, arrLog);

    if (rowOffsetSkills < tableMarksObtained.length && tableMarksObtained[rowOffsetSkills] != null) {

      var temp = tableMarksObtained[rowOffsetSkills].slice(1);

      if (temp != null) {

        for (j = 0; j < temp.length; j++) {

          if (temp[j] == null || temp[j].toString() == "") {

            arrSkills.push("not defined");

          } else {

            arrSkills.push(trimAndToLowerCaseAndRemoveExtraSpaces(temp[j]));
          }
        }
      }
    }

    for (i = 0; i < tableMarksObtained.length; i++) {

      studentId = null;
      isStudentIdAvailable = false;
      attendance = null
      isStudentPresent = false;
      arrMarksObtained = null;

      if (tableStudentDetails != null && i < tableStudentDetails.length) {

        studentId = tableStudentDetails[i][0];
        if (studentId != null && studentId.toString() != "") {

          isStudentIdAvailable = true;
        }
      }

      if (tableMarksObtained[i + rowNumHeadersInMarksObtained] != null &&
        tableMarksObtained[i + rowNumHeadersInMarksObtained].length > 0) {

        attendance = tableMarksObtained[i + rowNumHeadersInMarksObtained][0];
        if (attendance != null && attendance.toString().toLowerCase() == "p") isStudentPresent = true;

        arrMarksObtained = tableMarksObtained[i + rowNumHeadersInMarksObtained].slice(1);
      }

      if (arrMarksObtained != null) {

        for (j = 0; j < arrMarksObtained.length; j++) {

          dbRow = arrSubmissionDetails.slice();
          maxMarks = null;
          isMaxMarksValid = false;
          isMaxMarksEmpty = false;
          skill = "not defined";
          percentage = null;

          if (arrMaxMarks != null && j < arrMaxMarks.length) {

            maxMarks = arrMaxMarks[j];
          }

          if (j < arrIsMaxMarksValid.length) {

            isMaxMarksValid = arrIsMaxMarksValid[j];
          }

          if (j < arrIsMaxMarksEmpty.length) {

            isMaxMarksEmpty = arrIsMaxMarksEmpty[j];
          }

          marksObtained = arrMarksObtained[j];

          if (arrSkills != null && j < arrSkills.length) {

            skill = arrSkills[j];
          }

          isTrackerValid = isTrackerValid && validateRubricMarksObtained(listRubric, isMaxMarksEmpty, marksObtained, isStudentIdAvailable,
            isStudentPresent, i + rowOffsetMarksObtained, colOffsetMarksObtained, arrLog);


          if (isTrackerValid == true && isStudentIdAvailable == true && isMaxMarksValid == true) {

            marksObtained = rubricToMarks(listRubric, marksObtained);

            if (isStudentPresent == true) {

              percentage = marksObtained / maxMarks;
            }

            dbRow = dbRow.concat([studentId,
              tableStudentDetails[i][1],
              tableStudentDetails[i][6],
              tableStudentDetails[i][7],
              tableStudentDetails[i][8],
              tableStudentDetails[i][9],
              tableStudentDetails[i][10]],
              arrAssessmentDetails,
              [skill,
                attendance,
                maxMarks,
                marksObtained,
                percentage,
                "active"]);


            arrRowsToInsertInAssessmentDb.push(dbRow);
          }
        }

      } else if (isStudentIdAvailable == true) {

        arrLog.push("No marks entered for " + tableStudentDetails[i][1] + " (" + tableStudentDetails[studentId] + ")");
        isTrackerValid == false;
      }
    }
  }

  return isTrackerValid;
}


function validateRubricMaxMarks(maxRubricMarks, arrMaxMarks, arrIsMaxMarksValid, arrIsMaxMarksEmpty, colOffsetMarksObtained, rowTrackerMaxMarks, arrLog) {

  var i, maxMarks, typeMaxMarks, cellAddress, isTrackerValid = true;

  if (arrMaxMarks != null) {

    for (i = 0; i < arrMaxMarks.length; i++) {

      cellAddress = columnToLetter(i + colOffsetMarksObtained) + rowTrackerMaxMarks;

      maxMarks = arrMaxMarks[i];
      typeMaxMarks = typeof maxMarks;

      arrIsMaxMarksValid.push(false);
      arrIsMaxMarksEmpty.push(false);

      if (typeMaxMarks == (typeof 0)) {

        if (maxMarks == 0) {

          arrLog.push("Max Marks entered in " + cellAddress + " is zero");
          isTrackerValid = isTrackerValid && false;

        } else if (maxMarks < 0) {

          arrLog.push("Max Marks entered in " + cellAddress + " is negative");
          isTrackerValid = isTrackerValid && false;

        } else if (maxMarks != maxRubricMarks) {

          arrLog.push("Max Marks entered in " + cellAddress + " is not equal to the maximum rubric score");
          isTrackerValid = isTrackerValid && false;

        } else {

          arrIsMaxMarksValid[i] = true;
        }

      } else if (maxMarks == null || maxMarks.toString() == "") {

        arrIsMaxMarksEmpty[i] = true;

      } else {

        arrLog.push("Max Marks entered in " + cellAddress + " is not a valid number");
        isTrackerValid = isTrackerValid && false;
      }
    }
  }

  return isTrackerValid;
}


function validateRubricMarksObtained(listRubric, isMaxMarksEmpty, marksObtained, isStudentIdAvailable,
  isStudentPresent, rowNum, colNum, arrLog) {

  var i;
  var isValidRubricMark = false;
  var isTrackerValid = true;
  var cellAddress = columnToLetter(colNum) + rowNum;


  if (isStudentIdAvailable == true && isStudentPresent == true) {

    if (isMaxMarksEmpty == false) {

      for (i = 0; i < listRubric.length; i++) {

        if (listRubric[i][0].toString().toLowerCase() == marksObtained.toString().toLowerCase()) isValidRubricMark = true;
      }

      isTrackerValid = isTrackerValid && isValidRubricMark;

    } else if (marksObtained != null && marksObtained.toString() != "") {

      arrLog.push("There is either an extra entry in " + cellAddress + " or the Max Marks has not been entered in column " + columnToLetter(colNum));
      isTrackerValid = isTrackerValid && false;
    }

  } else {

    if (marksObtained != null && marksObtained.toString() != "") {

      arrLog.push("There is an extra entry in " + cellAddress + ". It should be empty");
      isTrackerValid = isTrackerValid && false;
    }
  }

  return isTrackerValid;
}

function rubricToMarks(listRubric, rubricMarksObtained) {

  var i;

  for (i = 0; i < listRubric.length; i++) {

    if (listRubric[i][0].toString().toLowerCase() == rubricMarksObtained.toString().toLowerCase()) {

      return i + 1;
    }
  }

  return 0;

}


function CreateGoogleSpreadSheet(userSubmittedTrackerId) {

  var userSubmittedTrackerFile = DriveApp.getFileById(userSubmittedTrackerId);
  userSubmittedTrackerFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var userSumbittedTrackerFileType = userSubmittedTrackerFile.getMimeType();

  var userSubmittedFileContents = userSubmittedTrackerFile.getBlob();

  const tempFolderId = "1RS03dDbf5d-f0MlnMNHxuyMHczYmWXVa";

  var resource = {
    title: userSubmittedTrackerFile.getName(),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: tempFolderId }]
  };

  var googleSheetWithUserSubmittedContent = Drive.Files.insert(resource, userSubmittedFileContents);
  var googleSheetWithUserSubmittedContentId = googleSheetWithUserSubmittedContent.id;

  return googleSheetWithUserSubmittedContentId;

}

function isValidDate(date, lockingPeriodInDays, log) {

  try {

    //var temp = Utilities.formatDate(date, "IST", "YYYY/MM/DD");
    
    var assessmentYear = Utilities.formatDate(date, "IST", "YYYY");
    var assessmentMonth = Utilities.formatDate(date, "IST", "M");
    var assessmentDayOfMonth = Utilities.formatDate(date, "IST", "d");

    var todayYear = Utilities.formatDate(new Date(), "IST", "YYYY");
    var todayMonth = Utilities.formatDate(new Date(), "IST", "M");
    var todayDayOfMonth = Utilities.formatDate(new Date(), "IST", "d");

    var dateDiff = Math.round((new Date(todayYear, todayMonth - 1, todayDayOfMonth) - new Date(assessmentYear, assessmentMonth - 1, assessmentDayOfMonth)) / (1000 * 60 * 60 * 24));

    if (dateDiff < 0) {

      log.push("The assessment date is in the future!");
      return false;

    } else if (dateDiff > lockingPeriodInDays) {

      log.push("You are submitting an assessment post the deadline of " + lockingPeriodInDays + " days. The tracker has not been accepted.");
      return false;

    }

  } catch (e) {

    if (date == null || date.toString() == "") {

      log.push("Date is empty");

    } else {

      log.push("Date entered is invalid");
    }
    return false;
  }

  return true;
}

function makeKey(componentArray) {

  var i, key, temp;

  if (componentArray != null && componentArray.length > 0) {

    key = trimAndToLowerCaseAndRemoveSpaces(componentArray[0]);

    for (i = 1; i < componentArray.length; i++) {

      key = key + "_" + trimAndToLowerCaseAndRemoveSpaces(componentArray[i]);
    }
  }

  return key;
}

function trimAndToLowerCaseAndRemoveSpaces(str) {

  if (str != null) {

    var temp = str.toString().replace(/^\s+|\s+$/g, '');
    temp = temp.replace(/\s+/g, '_');
    temp = temp.toLowerCase();

    return temp;
  }
}

function trimAndToLowerCaseAndRemoveExtraSpaces(str) {

  if (str != null) {

    var temp = str.replace(/^\s+|\s+$/g, '');
    temp = temp.replace(/\s+/g, ' ');
    temp = temp.toLowerCase();
    return temp;
  }
}

function letterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}