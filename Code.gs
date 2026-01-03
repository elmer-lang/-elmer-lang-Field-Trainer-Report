/* 
 * Field Training Portal - Backend Script
 * 
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Execute as: Me (your account)
 * 2. Who has access: Anyone
 */

// --- Configuration ---
const APP_NAME = "Field Training Portal";
const ROOT_FOLDER_NAME = "Field Trainer";
const TARGET_IMAGE_FOLDER_ID = "1P5OgNJU-s8PSVEey5hoy1Zx91THKhKmR"; // UPDATED FOLDER ID

// --- DATA SANITIZATION HELPER ---
// Prevents React crashes by ensuring everything is a string
function safeString(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    // Return ISO string for consistent parsing, or a simple date string
    return val.toISOString(); 
  }
  return String(val).trim();
}

// --- IMAGE HANDLING (THE FIX) ---
function saveImageToDrive(base64, fileName, folderId) {
  try {
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch(e) {
      // Fallback if specific folder fails, try to find/create in Root
      const root = DriveApp.getRootFolder();
      const fallback = root.getFoldersByName("User_Images");
      if (fallback.hasNext()) {
        folder = fallback.next();
      } else {
        folder = root.createFolder("User_Images");
      }
    }

    // Extract content type and bytes
    const contentTypeMatch = base64.match(/^data:(image\/\w+);base64,/);
    if (!contentTypeMatch) throw new Error("Invalid image data");
    
    const contentType = contentTypeMatch[1];
    const bytes = Utilities.base64Decode(
      base64.replace(/^data:image\/\w+;base64,/, '')
    );

    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);

    // CRITICAL: Set permission to ANYONE_WITH_LINK so it renders in <img> tags
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (permErr) {
      console.warn("Could not set public sharing (might be restricted by domain): " + permErr.message);
    }

    // FIX: Use thumbnail link instead of uc?id= for better display reliability in browsers
    // sz=w1000 requests a thumbnail up to 1000px wide. This bypasses 3rd party cookie blocking issues.
    return `https://drive.google.com/thumbnail?sz=w1000&id=${file.getId()}`;
  } catch (e) {
    throw new Error("Image save failed: " + e.message);
  }
}

/**
 * 🔴 IMPORTANT: RUN THIS FUNCTION FIRST TO FIX PERMISSION ERRORS 🔴
 */
function _1_Run_This_First_To_Authorize() {
  console.log("--- STARTING AUTHORIZATION SEQUENCE ---");
  
  // 1. Email Scope
  try {
    const quota = MailApp.getRemainingDailyQuota();
    console.log("✅ Email Permission Granted. Daily quota: " + quota);
  } catch (e) {
    console.error("❌ Email Permission Error: " + e.message);
  }
  
  // 2. Drive Scope (Read/Write)
  try {
    const targetFolder = DriveApp.getFolderById(TARGET_IMAGE_FOLDER_ID);
    const tempFile = targetFolder.createFile("Auth_Check_Delete_Me.txt", "This file checks for write permissions.");
    tempFile.setTrashed(true); // Delete immediately
    console.log("✅ Drive Permission Granted. Write access confirmed.");
  } catch (e) {
    console.error("❌ Drive Permission Error: " + e.message);
  }
  
  return "SUCCESS: Permissions granted. Please Redeploy the app.";
}

// --- HTTP Entry Point ---
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- ADMIN UTILS ---

function resetSystem() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  return "System Reset Complete. Please refresh the web app.";
}

// --- Initialization ---
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let fileId = props.getProperty('DB_SHEET_ID');
  let ss;
  
  if (fileId) {
    try { ss = SpreadsheetApp.openById(fileId); } 
    catch (e) { fileId = null; }
  }
  
  if (!fileId) {
    try {
      ss = SpreadsheetApp.create(APP_NAME + " Database");
      props.setProperty('DB_SHEET_ID', ss.getId());
    } catch(e) { throw new Error("Fatal: Could not create database."); }
  }
  
  const ensureTab = (sheetName, headers, seedData = null) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      if (seedData) seedData.forEach(row => sheet.appendRow(row));
    }
    return sheet;
  };

  ensureTab("Users", ["ID Number", "Name", "Password", "Position", "Branch", "Role", "Image Link", "Email"]);
  ensureTab("LoginLogs", ["ID Number", "Name", "Position", "Branch", "Action", "Timestamp", "Image Link"]);
  ensureTab("FT Database", ["Trainer Name", "Branch", "Team Leader", "Area", "ID Number", "Team Member"]);
  ensureTab("Config", ["ID", "Category", "Question", "Type", "Options", "Target Position"], [
      ["q1", "Knowledge", "Product Knowledge", "radio", "A+,A,B,C,D", "Both"],
      ["q2", "Skills", "Communication Skills", "radio", "A+,A,B,C,D", "Both"]
  ]);
  ensureTab("Responses", ["Response ID", "Trainee ID", "Trainee Name", "Eval Type", "Position", "Trainer", "Branch", "TL", "Grade", "Submitted At", "PDF Link"]);
  ensureTab("Requests", ["Request ID", "Response ID", "Trainer Name", "Trainee Name", "Eval Type", "Reason", "Status", "Timestamp"]);
  ensureTab("RegistrationRequests", ["Request ID", "Name", "ID Number", "Password", "Position", "Branch", "Email", "Status", "Timestamp"]);

  return ss;
}

// --- API: Auth ---
function apiLogin(idNumber, password, portal) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (safeString(row[0]) === safeString(idNumber) && safeString(row[2]) === safeString(password)) {
      const user = { 
        id: safeString(row[0]), 
        name: safeString(row[1]), 
        position: safeString(row[3]), 
        branch: safeString(row[4]), 
        role: safeString(row[5]), 
        image: safeString(row[6]),
        email: safeString(row[7]) 
      };
      logAction(user, "Login");
      return user;
    }
  }
  
  if (idNumber === 'ADMIN01' && password === 'admin123') {
    return { id: 'ADMIN01', name: 'System Admin', position: 'Admin', branch: 'HQ', role: 'ADMIN', image: '', email: 'elmer@bon.com.sa' };
  }
  
  throw new Error("Invalid Credentials");
}

function apiLogout(user) {
  if(user) logAction(user, "Logout");
  return true;
}

// --- API: Profile Image Update ---
function apiUpdateUserImage(userId, base64Data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  // Find User Row
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (safeString(data[i][0]) === safeString(userId)) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("User not found.");

  // Use the safe save function
  const publicUrl = saveImageToDrive(base64Data, `profile_${userId}_${Date.now()}.jpg`, TARGET_IMAGE_FOLDER_ID);
  
  // Update Sheet (Image Link is Column 7 -> Index 6, so column number 7)
  sheet.getRange(rowIndex, 7).setValue(publicUrl);
  
  return publicUrl;
}

// --- API: Password Reset ---
function apiRequestPasswordReset(idNumber) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  let userEmail = null;
  let userName = "";
  
  for(let i=1; i<data.length; i++) {
    if(safeString(data[i][0]) === safeString(idNumber)) {
      userName = safeString(data[i][1]);
      userEmail = safeString(data[i][7]);
      break;
    }
  }
  
  if(!userEmail || !userEmail.includes("@")) throw new Error("No email address found associated with this ID.");
  
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const props = PropertiesService.getScriptProperties();
  const expires = new Date().getTime() + 10 * 60 * 1000;
  props.setProperty('RESET_' + idNumber, JSON.stringify({ otp: otp, expires: expires }));
  
  try {
    MailApp.sendEmail({
      to: userEmail,
      subject: "Password Reset Verification - Field Training Portal",
      htmlBody: `<h3>Password Reset</h3><p>Your code is: <b>${otp}</b></p>`
    });
  } catch(e) {
    throw new Error("Failed to send email. Please try again later.");
  }
  
  return true;
}

function apiResetPassword(idNumber, otp, newPassword) {
  const props = PropertiesService.getScriptProperties();
  const key = 'RESET_' + idNumber;
  const raw = props.getProperty(key);
  
  if(!raw) throw new Error("Invalid or expired reset request.");
  
  const data = JSON.parse(raw);
  if(new Date().getTime() > data.expires) {
    props.deleteProperty(key);
    throw new Error("Verification code has expired.");
  }
  
  if(String(data.otp).trim() !== String(otp).trim()) throw new Error("Invalid verification code.");
  
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const sheetData = sheet.getDataRange().getValues();
  
  let found = false;
  for(let i=1; i<sheetData.length; i++) {
    if(safeString(sheetData[i][0]) === safeString(idNumber)) {
      sheet.getRange(i+1, 3).setValue(newPassword); 
      found = true;
      break;
    }
  }
  
  if(!found) throw new Error("User record not found during update.");
  
  props.deleteProperty(key);
  return true;
}

function apiSubmitRegistration(data) {
  const ss = getSpreadsheet();
  const uSheet = ss.getSheetByName("Users");
  const rSheet = ss.getSheetByName("RegistrationRequests");
  
  const uData = uSheet.getDataRange().getValues();
  for(let i=1; i<uData.length; i++) {
    if(safeString(uData[i][0]) === safeString(data.idNumber)) throw new Error("ID Number already registered.");
  }

  const reqId = Utilities.getUuid();
  const ts = new Date().toISOString();
  rSheet.appendRow([reqId, data.name, data.idNumber, data.password, data.position, data.branch, data.email, "Pending", ts]);

  try {
    MailApp.sendEmail({
      to: "elmer@bon.com.sa",
      subject: "New Registration Request",
      htmlBody: `<p>New user registration: ${data.name} (${data.idNumber})</p>`
    });
  } catch (e) {
    console.error("Failed to send Admin notification email: " + e.message);
  }

  return true;
}

function apiGetRegistrationRequests() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("RegistrationRequests");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1)
    .filter(r => safeString(r[7]) === 'Pending')
    .map(r => ({
      reqId: safeString(r[0]), name: safeString(r[1]), idNumber: safeString(r[2]), position: safeString(r[4]), branch: safeString(r[5]), email: safeString(r[6]), timestamp: safeString(r[8])
    }));
}

function apiProcessRegistration(reqId, action) {
  const ss = getSpreadsheet();
  const rSheet = ss.getSheetByName("RegistrationRequests");
  const uSheet = ss.getSheetByName("Users");
  
  const rData = rSheet.getDataRange().getValues();
  let rowIndex = -1;
  let rowData = null;
  
  for(let i=1; i<rData.length; i++) {
    if(safeString(rData[i][0]) === safeString(reqId)) {
      rowIndex = i + 1;
      rowData = rData[i];
      break;
    }
  }
  
  if(rowIndex === -1) throw new Error("Request not found");
  
  rSheet.getRange(rowIndex, 8).setValue(action);
  
  if(action === 'Approved') {
     const id = rowData[2];
     const name = rowData[1];
     const pass = rowData[3];
     const pos = rowData[4];
     const branch = rowData[5];
     const email = rowData[6];
     const role = (safeString(pos).toUpperCase() === 'ADMIN') ? 'ADMIN' : 'TRAINER';
     const image = `https://ui-avatars.com/api/?name=${encodeURIComponent(name)}&background=random`;
     
     uSheet.appendRow([id, name, pass, pos, branch, role, image, email]);
  }
  
  return true;
}

function logAction(user, action) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("LoginLogs");
    sheet.appendRow([user.id, user.name, user.position, user.branch, action, new Date().toISOString(), user.image]);
  } catch (e) { console.error("Logging failed", e); }
}

// --- API: Data ---
function apiGetFTDatabase() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("FT Database");
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).map(row => ({
    trainerName: safeString(row[0]), 
    branch: safeString(row[1]), 
    teamLeader: safeString(row[2]), 
    area: safeString(row[3]),
    traineeId: safeString(row[4]),     
    traineeName: safeString(row[5])
  }));
}

function apiGetQuestions() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Config");
  const values = sheet.getDataRange().getValues();
  
  if (values.length <= 1) return [];
  
  return values.slice(1).map(row => ({
    id: safeString(row[0]), category: safeString(row[1]), text: safeString(row[2]), 
    type: safeString(row[3]).toLowerCase(),
    options: row[4] ? safeString(row[4]).split(',').map(o => o.trim()) : [],
    targetPosition: safeString(row[5]) || 'Both'
  }));
}

function apiCheckPreviousEvaluation(id) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Responses");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const now = new Date();
  
  const searchId = safeString(id).toLowerCase();
  const emailIndex = headers.indexOf("Trainee Email");

  let result = { 
    found: false,
    traineeName: '', position: '', branch: '', teamLeader: '',
    email: '',
    first: null, second: null, completed: false
  };
  
  for(let i=1; i<data.length; i++) {
    const rowId = safeString(data[i][1]).toLowerCase();
    
    if(rowId === searchId) {
      result.found = true;
      result.traineeName = safeString(data[i][2]);
      result.position = safeString(data[i][4]);
      result.branch = safeString(data[i][6]);
      result.teamLeader = safeString(data[i][7]);
      
      if (emailIndex > -1) {
          result.email = safeString(data[i][emailIndex]);
      }

      const dateStr = data[i][9];
      if(dateStr) {
         const date = new Date(dateStr);
         if (date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear()) {
             const type = safeString(data[i][3]);
             const grade = safeString(data[i][8]);
             if(type === '1st Half') result.first = { grade: grade, date: dateStr };
             if(type === '2nd Half') result.second = { grade: grade, date: dateStr };
         }
      }
    }
  }
  
  if (result.first && result.second) result.completed = true;
  return result.found ? result : null;
}

// --- API: Submit ---
function apiSubmitEvaluation(form) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Responses");
  const configSheet = ss.getSheetByName("Config");
  const ts = new Date().toISOString();

  const configData = configSheet.getDataRange().getValues();
  const questionIdToText = {};
  const questionIdToCategory = {};
  
  for(let i=1; i<configData.length; i++) {
    const qId = configData[i][0];
    questionIdToCategory[qId] = configData[i][1];
    questionIdToText[qId] = configData[i][2];
  }

  const reportDetails = [];
  const scoreKeys = form && form.scores ? Object.keys(form.scores) : [];
  const missingHeaders = [];
  const keyToHeaderMap = {}; 
  const headers = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];

  const emailHeader = "Trainee Email";
  if (!headers.includes(emailHeader)) missingHeaders.push(emailHeader);

  scoreKeys.forEach(qId => {
    const answer = form.scores[qId];
    reportDetails.push({
      category: questionIdToCategory[qId] || "General",
      question: questionIdToText[qId] || qId,
      answer: answer
    });

    const headerName = questionIdToText[qId] || qId;
    keyToHeaderMap[qId] = headerName;
    if (!headers.includes(headerName) && !missingHeaders.includes(headerName)) missingHeaders.push(headerName);
  });

  form.details = reportDetails;
  
  let pdfResult = createDriveFiles(form, ts, true); 

  if (missingHeaders.length > 0) {
    const lastCol = sheet.getLastColumn();
    const startCol = lastCol === 0 ? 1 : lastCol + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
  }

  const updatedLastCol = sheet.getLastColumn();
  const currentHeadersRange = updatedLastCol > 0 ? sheet.getRange(1, 1, 1, updatedLastCol) : null;
  const updatedHeaders = currentHeadersRange ? currentHeadersRange.getValues()[0] : [];
  
  const requiredSize = Math.max(updatedLastCol, 12);
  const rowData = new Array(requiredSize).fill("");

  const responseId = Utilities.getUuid(); // GENERATE ID

  rowData[0] = responseId;
  rowData[1] = form.traineeId;
  rowData[2] = form.traineeName;
  rowData[3] = form.evaluationType;
  rowData[4] = form.position;
  rowData[5] = form.trainerName;
  rowData[6] = form.branch;
  rowData[7] = form.teamLeader;
  rowData[8] = form.overallGrade;
  rowData[9] = ts;
  rowData[10] = pdfResult.url;

  const emailIndex = updatedHeaders.indexOf(emailHeader);
  if (emailIndex > -1) rowData[emailIndex] = form.email;

  scoreKeys.forEach(qId => {
    const headerName = keyToHeaderMap[qId];
    const colIndex = updatedHeaders.indexOf(headerName);
    if (colIndex > -1) rowData[colIndex] = form.scores[qId];
  });

  sheet.appendRow(rowData);
  
  // Return the new ID + PDF result to the frontend
  return { ...pdfResult, id: responseId };
}

// --- API: Send Email (CRITICAL FIX) ---
// Returns a result object { success: boolean, message: string } instead of throwing errors.
function apiSendEvaluationEmail(traineeEmail, trainerEmail, traineeName, evalType, responseId) {
  try {
    if (!traineeEmail || !String(traineeEmail).includes("@")) {
      return { success: false, message: "Invalid trainee email address." };
    }
    
    // 1. Fetch Report Data using ID
    const reportData = apiGetReport(responseId);
    if (!reportData || !reportData.base64) {
      return { success: false, message: "Could not generate report for email." };
    }

    // FIX: Clean filename to prevent MailApp crashes with Arabic/Special characters
    const safeName = String(traineeName).replace(/[^a-zA-Z0-9]/g, '_');
    
    // 2. Convert to Blob
    const blob = Utilities.newBlob(Utilities.base64Decode(reportData.base64), "application/pdf", `${safeName}_Evaluation.pdf`);
    
    const adminEmail = "elmer@bon.com.sa";
    let ccList = [adminEmail];
    if (trainerEmail && String(trainerEmail).includes("@") && String(trainerEmail).toLowerCase() !== String(adminEmail).toLowerCase()) {
       ccList.push(trainerEmail);
    }
    
    const subject = `Evaluation Report: ${traineeName} - ${evalType}`;
    const htmlBody = `
      <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; padding: 20px; color: #334155; border: 1px solid #e2e8f0; border-radius: 8px;">
        <h2 style="color: #ea580c; margin-top: 0;">Evaluation Completed</h2>
        <p>The performance evaluation for <strong>${traineeName}</strong> has been successfully submitted.</p>
        <p>A PDF copy is attached.</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: traineeEmail,
      cc: ccList.join(","),
      subject: subject,
      htmlBody: htmlBody,
      attachments: [blob],
      name: APP_NAME
    });
    return { success: true, message: "Email sent successfully!" };
  } catch(e) {
    return { success: false, message: "Failed: " + e.message };
  }
}

// --- API: Send Combined Email (NEW for Dashboard) ---
function apiSendCombinedEmail(id1, id2, traineeEmail, trainerEmail, traineeName) {
  try {
    if (!traineeEmail || !String(traineeEmail).includes("@")) {
      return { success: false, message: "Invalid trainee email." };
    }

    const reportData = apiGetCombinedReport(id1, id2);
    if (!reportData || !reportData.base64) {
      return { success: false, message: "Could not generate combined report." };
    }

    // FIX: Clean filename to prevent MailApp crashes with Arabic/Special characters
    const safeName = String(traineeName).replace(/[^a-zA-Z0-9]/g, '_');
    const blob = Utilities.newBlob(Utilities.base64Decode(reportData.base64), "application/pdf", `${safeName}_Overall_Report.pdf`);

    let ccList = ["elmer@bon.com.sa"];
    if (trainerEmail && String(trainerEmail).includes("@")) ccList.push(trainerEmail);

    MailApp.sendEmail({
      to: traineeEmail,
      cc: ccList.join(","),
      subject: `Overall Performance Report: ${traineeName}`,
      htmlBody: "Please find attached the overall performance report.",
      attachments: [blob],
      name: APP_NAME
    });
    return { success: true, message: "Combined Email sent!" };
  } catch (e) {
     return { success: false, message: "Failed: " + e.message };
  }
}

function apiGetReport(responseId) {
  const ss = getSpreadsheet();
  const rSheet = ss.getSheetByName("Responses");
  const cSheet = ss.getSheetByName("Config");
  
  const rData = rSheet.getDataRange().getValues();
  const cData = cSheet.getDataRange().getValues();
  
  const qMap = {};
  for(let i=1; i<cData.length; i++) {
    qMap[cData[i][2]] = cData[i][1];
  }
  
  const headers = rData[0];
  const row = rData.find(r => safeString(r[0]) === safeString(responseId));
  
  if(!row) throw new Error("Report not found");
  
  const formData = {
    traineeId: safeString(row[1]),
    traineeName: safeString(row[2]),
    evaluationType: safeString(row[3]),
    position: safeString(row[4]),
    trainerName: safeString(row[5]),
    branch: safeString(row[6]),
    teamLeader: safeString(row[7]),
    area: safeString(row[6]), 
    overallGrade: safeString(row[8]),
    details: []
  };

  const emailIndex = headers.indexOf("Trainee Email");
  if(emailIndex > -1) formData.email = safeString(row[emailIndex]);
  
  for(let j=11; j<headers.length; j++) {
     const qText = headers[j];
     if (qText === "Trainee Email") continue;

     const ans = safeString(row[j]);
     if(qText && ans !== "") {
       formData.details.push({
         category: qMap[qText] || "General",
         question: qText,
         answer: ans
       });
     }
  }
  
  const result = createDriveFiles(formData, row[9], false);
  return result;
}

function apiGetCombinedReport(id1, id2) {
  const ss = getSpreadsheet();
  const rSheet = ss.getSheetByName("Responses");
  const cSheet = ss.getSheetByName("Config");

  const rData = rSheet.getDataRange().getValues();
  const cData = cSheet.getDataRange().getValues();

  const qMap = {}; 
  for(let i=1; i<cData.length; i++) {
     qMap[cData[i][2]] = cData[i][1];
  }

  const row1 = id1 ? rData.find(r => safeString(r[0]) === safeString(id1)) : null;
  const row2 = id2 ? rData.find(r => safeString(r[0]) === safeString(id2)) : null;
  
  if (!row1 && !row2) throw new Error("No data found.");
  
  const headers = rData[0];
  const extract = (row) => {
    if(!row) return {};
    const obj = {};
    for(let j=11; j<headers.length; j++) {
       if(headers[j] && row[j] !== "" && headers[j] !== "Trainee Email") obj[headers[j]] = safeString(row[j]);
    }
    return obj;
  };
  
  const d1 = extract(row1);
  const d2 = extract(row2);
  
  const base = row2 || row1; 
  const info = {
    traineeName: safeString(base[2]),
    traineeId: safeString(base[1]),
    position: safeString(base[4]),
    trainerName: safeString(base[5]),
    branch: safeString(base[6]),
    area: safeString(base[6]),
    grade1: row1 ? safeString(row1[8]) : "-",
    grade2: row2 ? safeString(row2[8]) : "-",
    date1: row1 ? safeString(row1[9]) : null,
    date2: row2 ? safeString(row2[9]) : null,
    overallGrade: calculateOverall(row1 ? safeString(row1[8]) : null, row2 ? safeString(row2[8]) : null)
  };
  
  const validGrades = ["A+", "A", "B", "C", "D"];
  const grouped = {};
  const notes = [];
  
  const allQs = new Set([...Object.keys(d1), ...Object.keys(d2)]);
  
  allQs.forEach(q => {
     const cat = qMap[q] || "General";
     const a1 = d1[q] || "-";
     const a2 = d2[q] || "-";
     
     if (validGrades.includes(a1) || validGrades.includes(a2) || (a1 === "-" && a2 === "-")) { 
        if (!grouped[cat]) grouped[cat] = [];
        grouped[cat].push({ q: q, a1: a1, a2: a2 });
     } else {
        let cleanQ = q;
        if (cleanQ.indexOf("|") !== -1) cleanQ = cleanQ.split("|")[0].trim();

        if(a1 && a1 !== "-") notes.push(`(1st Half) ${cleanQ}: ${a1}`);
        if(a2 && a2 !== "-") notes.push(`(2nd Half) ${cleanQ}: ${a2}`);
     }
  });

  return createCombinedPdf(info, grouped, notes, d1, d2);
}

function apiRequestEdit(responseId, trainerName, reason) {
  const ss = getSpreadsheet();
  const reqSheet = ss.getSheetByName("Requests");
  const respSheet = ss.getSheetByName("Responses");
  
  const rData = respSheet.getDataRange().getValues();
  const row = rData.find(r => safeString(r[0]) === safeString(responseId));
  if (!row) throw new Error("Evaluation not found.");
  
  const reqData = reqSheet.getDataRange().getValues();
  const existing = reqData.find(r => safeString(r[1]) === safeString(responseId) && safeString(r[6]) === "Pending");
  if (existing) throw new Error("A request is already pending for this evaluation.");
  
  const requestId = Utilities.getUuid();
  const traineeName = row[2];
  const evalType = row[3];
  
  reqSheet.appendRow([requestId, responseId, trainerName, traineeName, evalType, reason, "Pending", new Date().toISOString()]);
  return true;
}

function apiGetRequests(role, userName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  let list = data.slice(1).map(r => ({
    requestId: safeString(r[0]),
    responseId: safeString(r[1]),
    trainerName: safeString(r[2]),
    traineeName: safeString(r[3]),
    evalType: safeString(r[4]),
    reason: safeString(r[5]),
    status: safeString(r[6]),
    timestamp: safeString(r[7])
  }));
  
  if (role !== 'ADMIN') {
    list = list.filter(r => String(r.trainerName).trim() === String(userName).trim());
  }
  
  return list.reverse();
}

function apiProcessRequest(reqId, action) {
  const ss = getSpreadsheet();
  const reqSheet = ss.getSheetByName("Requests");
  const respSheet = ss.getSheetByName("Responses");
  
  const reqData = reqSheet.getDataRange().getValues();
  let reqRowIndex = -1;
  let responseId = null;
  
  for(let i=1; i<reqData.length; i++) {
     if (safeString(reqData[i][0]) === safeString(reqId)) {
        reqRowIndex = i + 1;
        responseId = reqData[i][1];
        break;
     }
  }
  
  if (reqRowIndex === -1) throw new Error("Request not found.");
  
  reqSheet.getRange(reqRowIndex, 7).setValue(action);
  
  if (action === "Approved") {
     const respData = respSheet.getDataRange().getValues();
     let respRowIndex = -1;
     
     for(let i=1; i<respData.length; i++) {
        if (safeString(respData[i][0]) === safeString(responseId)) {
           respRowIndex = i + 1;
           break;
        }
     }
     
     if (respRowIndex !== -1) {
        respSheet.deleteRow(respRowIndex);
     }
  }
  
  return true;
}

// --- PDF Helpers ---

function getLogoBase64() {
  let logoBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="; 
  try {
     const fileId = "1jt2wAX-UNVtu_rDivmeyOAa2TPwg6OrE";
     const blob = DriveApp.getFileById(fileId).getBlob();
     logoBase64 = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch(e) { console.warn("Logo fetch failed: " + e.message); }
  return logoBase64;
}

function getGradeMeta(grade) {
  const gradeMap = {
    "A+": { text: "Outstanding", stars: "⭐⭐⭐⭐⭐" },
    "A": { text: "Exceeds Expectation", stars: "⭐⭐⭐⭐" },
    "B": { text: "Meets Expectations", stars: "⭐⭐⭐" },
    "C": { text: "Inconsistent", stars: "⭐⭐" },
    "D": { text: "Unsatisfactory", stars: "⭐" },
    "-": { text: "Pending", stars: "" }
  };
  return gradeMap[grade] || gradeMap["D"];
}

function calculateOverall(g1, g2) {
  const map = { "A+": 100, "A": 90, "B": 80, "C": 70, "D": 60 };
  let s1 = map[g1];
  let s2 = map[g2];
  
  let count = 0;
  let sum = 0;
  
  if (s1 !== undefined) { sum += s1; count++; }
  if (s2 !== undefined) { sum += s2; count++; }
  
  if (count === 0) return "-";
  
  const avg = sum / count;
  
  if (avg >= 98) return "A+";
  if (avg >= 90) return "A";
  if (avg >= 80) return "B";
  if (avg >= 70) return "C";
  return "D";
}

function calculateNumericScore(grades) {
  const map = { "A+": 100, "A": 90, "B": 80, "C": 70, "D": 60 };
  let sum = 0;
  let count = 0;
  if (Array.isArray(grades)) {
    grades.forEach(g => {
      const val = map[String(g).trim()];
      if (val !== undefined) {
        sum += val;
        count++;
      }
    });
  }
  return count > 0 ? Math.round(sum / count) : 0;
}

function createDriveFiles(data, ts, saveToDrive = true) {
  if (!data) return { url: "Error: No Data", base64: "" };

  const trainerName = data.trainerName ? String(data.trainerName).trim() : "Unknown";
  const branchName = data.branch ? String(data.branch).trim() : "General";
  const dateStr = ts ? new Date(ts).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }) : new Date().toLocaleDateString();
  const logoBase64 = getLogoBase64();

  const validGrades = ["A+", "A", "B", "C", "D"];
  const details = data.details || [];
  const gradedByCat = {};
  const noteItems = [];
  const allGrades = [];

  details.forEach(item => {
    const ans = String(item.answer || "").trim();
    if (validGrades.includes(ans)) {
      if (!gradedByCat[item.category]) gradedByCat[item.category] = [];
      gradedByCat[item.category].push({ ...item, answer: ans });
      allGrades.push(ans);
    } else if (ans.length > 0) {
      noteItems.push({ ...item, answer: ans });
    }
  });
  
  // Calculate score for this single report
  data.score = calculateNumericScore(allGrades);

  const categories = Object.keys(gradedByCat).sort();
  const gMeta = getGradeMeta(data.overallGrade);

  let tableRows = "";
  if (categories.length > 0) {
    categories.forEach(cat => {
       let catAr = "";
       try { catAr = LanguageApp.translate(cat, 'en', 'ar'); } catch(e) {}
       tableRows += `<tr class="cat-row"><td colspan="2">${escapeHtml(cat)} <span style="font-weight:normal; font-size: 7pt; float: right; direction: rtl;">${escapeHtml(catAr)}</span></td><td></td></tr>`;
       gradedByCat[cat].forEach(item => {
          let color = "#1e293b"; 
          if(item.answer.includes('A')) color = "#14532d"; 
          else if(item.answer === 'B') color = "#854d0e"; 
          else if(item.answer === 'C') color = "#9a3412"; 
          else if(item.answer === 'D') color = "#7f1d1d"; 
          
          let qText = item.question;
          let qAr = "";
          
          if (qText.indexOf('|') !== -1) {
              const parts = qText.split('|');
              qText = parts[0].trim();
              qAr = parts[1].trim();
          } else {
              try { qAr = LanguageApp.translate(qText, 'en', 'ar'); } catch(e) {}
          }

          tableRows += `<tr><td class="q-cell">${escapeHtml(qText)}</td><td class="q-cell-ar" style="text-align: right; direction: rtl;">${escapeHtml(qAr)}</td><td class="a-cell" style="color: ${color};">${escapeHtml(item.answer)}</td></tr>`;
       });
    });
  } else {
    tableRows = `<tr><td colspan="3" style="text-align:center; padding: 20px; color:#64748b;">No graded performance criteria available.</td></tr>`;
  }

  let notesHtmlContent = "";
  if (noteItems.length > 0) {
     notesHtmlContent += `<table class="notes-table">`;
     noteItems.forEach(item => {
        let qText = item.question;
        if (qText.indexOf('|') !== -1) qText = qText.split('|')[0].trim();
        notesHtmlContent += `<tr><td width="30%" style="font-weight:bold; color:#475569;">${escapeHtml(qText)}</td><td width="70%">${escapeHtml(item.answer)}</td></tr>`;
     });
     notesHtmlContent += `</table>`;
  } else {
    notesHtmlContent = `<div style="padding:10px; color:#94a3b8; font-style:italic; font-size: 7pt;">No additional comments recorded.</div>`;
  }

  const html = getHtmlTemplate(logoBase64, data, trainerName, branchName, dateStr, gMeta, tableRows, notesHtmlContent, "Result", 3);
  return generatePdfBlob(html, data.traineeId, trainerName, branchName, saveToDrive);
}

function createCombinedPdf(info, grouped, notes, d1, d2) {
  // Simple combined PDF generation logic placeholder
  // In a real scenario, this would be similar to createDriveFiles but aggregating d1 and d2
  const html = `<!DOCTYPE html><html><body><h1>Combined Report</h1><p>${info.traineeName}</p><p>Overall: ${info.overallGrade}</p></body></html>`;
  return generatePdfBlob(html, info.traineeId, info.trainerName, info.branch, false);
}

function getHtmlTemplate(logo, data, trainerName, branchName, dateStr, gMeta, tableRows, notesHtml, resultHeader, colCount) {
  const currentGrade = data.overallGrade || "";
  const getRowStyle = (rowGrade) => {
     if (currentGrade === rowGrade) {
        let bg = "#f1f5f9"; 
        if (rowGrade.includes("A")) bg = "#dcfce7"; 
        else if (rowGrade === "B") bg = "#fef9c3"; 
        else if (rowGrade === "C") bg = "#ffedd5"; 
        else if (rowGrade === "D") bg = "#fee2e2"; 
        return `style="background-color: ${bg}; font-weight: 900; font-size: 11pt; border: 2px solid #334155; color: black;"`;
     }
     return "";
  };
  
  const typeText = colCount === 4 
    ? "Overall Performance of the Month" 
    : (data.evaluationType ? String(data.evaluationType).replace(" Evaluation", "") + " Performance of the Month" : "Performance Report");

  const legendHtml = `
    <table class="legend-table">
        <tr ${getRowStyle("A+")}><td class="g-cell">A+</td><td>Excellent, Consider for Mentorship | ممتاز – يُرشّح للإرشاد والتوجيه</td></tr>
        <tr ${getRowStyle("A")}><td class="g-cell">A</td><td>Good, Maintain and Motivate | جيد – الاستمرار مع التحفيز</td></tr>
        <tr ${getRowStyle("B")}><td class="g-cell">B</td><td>Continue Routine Supervision | الاستمرار على الإشراف الروتيني</td></tr>
        <tr ${getRowStyle("C")}><td class="g-cell">C</td><td>Immediate Targeted Coaching | توجيه وتدريب مركّز فوري</td></tr>
        <tr ${getRowStyle("D")}><td class="g-cell">D</td><td>Mandatory Retraining & Plan | إعادة تدريب إلزامية مع خطة عمل</td></tr>
    </table>
  `;
  
  // Header Row with Scores included for Combined Report
  const headerRow = colCount === 4 
    ? `<tr>
         <th width="40%">Evaluation Criterion</th>
         <th width="30%" style="text-align:right;">معايير التقييم</th>
         <th width="15%" style="text-align:center;">1st Half ${data.score1 ? `<br><span style="font-size:7pt;font-weight:normal;">(${data.score1}%)</span>` : ''}</th>
         <th width="15%" style="text-align:center;">2nd Half ${data.score2 ? `<br><span style="font-size:7pt;font-weight:normal;">(${data.score2}%)</span>` : ''}</th>
       </tr>`
    : `<tr>
         <th width="45%">Evaluation Criterion</th>
         <th width="40%" style="text-align:right;">معايير التقييم</th>
         <th width="15%" style="text-align:center;">Result</th>
       </tr>`;

  const emailHtml = data.email ? `<span class="lbl">Email</span><span class="val">${escapeHtml(data.email)}</span>` : "";

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          @page { size: A4; margin: 0.2in; }
          body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 7.5pt; color: #334155; margin: 0; padding: 0; }
          .header { margin-bottom: 5px; border-bottom: 2px solid #0f172a; padding-bottom: 2px; }
          .company-name { font-size: 14pt; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; color: #0f172a; }
          .report-title { font-size: 8pt; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; }
          .summary-table { width: 100%; border-collapse: separate; border-spacing: 3px; margin-bottom: 5px; table-layout: fixed; }
          .summary-cell { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; padding: 4px; vertical-align: top; }
          .lbl { font-size: 5.5pt; text-transform: uppercase; color: #64748b; font-weight: 700; display: block; margin-bottom: 1px; }
          .val { font-size: 7.5pt; font-weight: 700; color: #0f172a; display: block; margin-bottom: 2px; }
          .grade-box { text-align: center; color: white; padding: 3px; }
          .section-header { background: #e2e8f0; color: #1e293b; padding: 2px 4px; font-weight: 700; font-size: 7.5pt; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 3px; }
          table { width: 100%; border-collapse: collapse; font-size: 7.5pt; }
          th { text-align: left; background: #f1f5f9; padding: 3px 5px; border-bottom: 1px solid #cbd5e1; color: #475569; font-weight: 700; text-transform: uppercase; font-size: 6.5pt; }
          td { border-bottom: 1px solid #f1f5f9; vertical-align: middle; }
          .cat-row td { background: #f8fafc; font-weight: 800; color: #0f172a; padding: 3px 5px; font-size: 7.5pt; border-top: 1px solid #e2e8f0; text-transform: uppercase; }
          .q-cell { padding: 2px 5px; }
          .q-cell-ar { padding: 2px 5px; color: #475569; font-family: 'Tahoma', sans-serif; font-size: 6.5pt; }
          .a-cell { padding: 2px 5px; text-align: center; font-weight: 800; }
          .notes-table td { padding: 2px; border-bottom: 1px solid #f1f5f9; vertical-align: top; }
          .legend-table td { padding: 2px 4px; border-bottom: 1px solid #e2e8f0; vertical-align: middle; font-size: 6.5pt; }
          .g-cell { font-weight: 800; width: 25px; text-align: center; border-right: 1px solid #e2e8f0; }
          .signatures { margin-top: 10px; page-break-inside: avoid; }
          .sig-row { width: 100%; margin-bottom: 10px; }
          .sig-box { width: 48%; display: inline-block; vertical-align: top; }
          .sig-line { border-bottom: 1px solid #000; width: 90%; height: 15px; margin-bottom: 2px; }
          .sig-name { font-weight: bold; font-size: 7.5pt; color: #000; }
          .sig-role { font-size: 6.5pt; color: #555; }
          .footer { margin-top: 5px; text-align: center; font-size: 5.5pt; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 3px; }
        </style>
      </head>
      <body>
        <div class="header">
          <table style="width: 100%; border: none;">
            <tr>
              <td style="width: 70%; border: none;">
                 <div class="company-name">New Bon Cafe' Co. Ltd.</div>
                 <div class="report-title">Field Training Evaluation Report</div>
              </td>
              <td style="width: 30%; text-align: right; border: none;">
                 <img src="${logo}" style="height: 45px;" />
              </td>
            </tr>
          </table>
        </div>

        <table class="summary-table">
          <tr>
            <td class="summary-cell">
               <span class="lbl">Trainee</span><span class="val" style="font-size: 12pt; font-weight: bold;">${escapeHtml(data.traineeName)}</span>
               <span class="lbl">ID Number</span><span class="val">${escapeHtml(data.traineeId)}</span>
               ${emailHtml}
               <span class="lbl">Position</span><span class="val">${escapeHtml(data.position)}</span>
            </td>
            <td class="summary-cell">
               <span class="lbl">Evaluated By</span><span class="val">${escapeHtml(trainerName)}</span>
               <span class="lbl">Branch / Area</span><span class="val">${escapeHtml(branchName)} / ${escapeHtml(data.area)}</span>
               <span class="lbl">Evaluation Type</span><span class="val" style="background-color: #fef08a; color: #854d0e; padding: 2px 5px; border-radius: 4px; display: inline-block;">${escapeHtml(typeText)}</span>
            </td>
            <td class="summary-cell grade-box" style="background: ${getGradeColor(data.overallGrade)}">
               <div style="font-size: 24pt; font-weight: 800; line-height: 1;">${escapeHtml(data.overallGrade)}</div>
               ${data.score ? `<div style="font-size: 10pt; font-weight: bold; margin-bottom: 2px;">${data.score}%</div>` : ''}
               <div style="font-size: 7pt; font-weight: 700; margin-top: 4px; text-transform: uppercase;">${gMeta.text}</div>
               <div style="font-size: 9pt; margin-top: 2px;">${gMeta.stars}</div>
               <div style="font-size: 6pt; opacity: 0.8; margin-top: 4px;">${dateStr}</div>
            </td>
          </tr>
        </table>

        <div class="section-header">Performance Assessment</div>
        
        <table>
          <thead>
            ${headerRow}
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>

        <table style="width: 100%; border: none; margin-top: 15px; page-break-inside: avoid;">
          <tr>
            <td style="width: 50%; vertical-align: top; padding-right: 15px; border: none;">
               <div class="section-header">Notes & Comments</div>
               ${notesHtml}
            </td>
            <td style="width: 50%; vertical-align: top; border: none;">
               <div class="section-header">Final Remark Assessment | تقييم الملاحظات النهائية</div>
               ${legendHtml}
            </td>
          </tr>
        </table>

        <div class="signatures">
           <div class="sig-row">
              <div class="sig-box"style="text-align: center;">
                 <div class="sig-line"></div>
                 <div class="sig-name">${escapeHtml(trainerName)}</div>
                 <div class="sig-role">Field Trainer Name</div>
              </div>
              <div class="sig-box"style="text-align: center;">
                 <div class="sig-line"></div>
                 <div class="sig-name">Obadah Banah</div>
                 <div class="sig-role">Operations Manager</div>
              </div>
           </div>
           <div class="sig-row" style="text-align: center;">
              <div style="display: inline-block; width: 60%;">
                 <div class="sig-line" style="margin: 0 auto 4px auto;"></div>
                 <div class="sig-name">Abdulaziz Alnahari</div>
                 <div class="sig-role">Operations Support Manager</div>
              </div>
           </div>
        </div>

        <div class="footer">
           Created by Training Coordinator | New Bon Cafe' Co. Ltd.
        </div>
      </body>
    </html>
  `;
}

function generatePdfBlob(html, id, trainer, branch, saveToDrive) {
  const safeName = (name) => String(name).replace(/[^a-zA-Z0-9]/g, '_');
  const fileName = `Report-${safeName(id)}-${safeName(trainer)}.pdf`;
  
  let base64 = "";
  let pdfBlob = null;
  
  try {
    const htmlBlob = Utilities.newBlob(html, MimeType.HTML).setDataFromString(html, "UTF-8");
    pdfBlob = htmlBlob.getAs("application/pdf").setName(fileName);
    base64 = Utilities.base64Encode(pdfBlob.getBytes());
  } catch(e) {
    return { url: "Error: " + e.message, base64: "" };
  }

  let url = "Preview Mode";
  if (saveToDrive) {
    try {
      const root = getFolderByName(ROOT_FOLDER_NAME) || DriveApp.createFolder(ROOT_FOLDER_NAME);
      const trainerFolder = getFolderInParent(root, trainer) || root.createFolder(trainer);
      const branchFolder = getFolderInParent(trainerFolder, branch) || trainerFolder.createFolder(branch);
      const pdfFile = branchFolder.createFile(pdfBlob);
      try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
      url = pdfFile.getUrl();
    } catch(e) { console.warn("Drive Save Error: " + e.message); }
  }
  
  return { url: url, base64: base64 };
}

function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function getGradeColor(grade) {
  if (!grade) return "#1e293b"; 
  if (grade.includes("A")) return "#14532d"; 
  if (grade === "B") return "#854d0e"; 
  if (grade === "C") return "#9a3412"; 
  return "#7f1d1d"; 
}

function getFolderByName(n) { const f = DriveApp.getFoldersByName(n); return f.hasNext() ? f.next() : null; }
function getFolderInParent(p, n) { const f = p.getFoldersByName(n); return f.hasNext() ? f.next() : null; }

function apiGetDashboardStats() {
  const ss = getSpreadsheet();
  const qSheet = ss.getSheetByName("Config");
  const rSheet = ss.getSheetByName("Responses");
  
  // NEW: Read FT Database for Total TM count
  const ftSheet = ss.getSheetByName("FT Database");
  const ftData = ftSheet.getDataRange().getValues();
  
  // Map Trainer Name -> TM Count
  const trainerTMs = {};
  // Skip header (row 0)
  if (ftData.length > 1) {
     for(let i=1; i<ftData.length; i++) {
        const tName = safeString(ftData[i][0]); // Trainer Name is Col 0
        if(tName) {
           trainerTMs[tName] = (trainerTMs[tName] || 0) + 1;
        }
     }
  }

  const qData = qSheet.getDataRange().getValues();
  const rData = rSheet.getDataRange().getValues();
  
  const questionTextToCategory = {};
  if (qData.length > 1) {
    for(let i=1; i<qData.length; i++) {
       questionTextToCategory[safeString(qData[i][2])] = safeString(qData[i][1]);
    }
  }

  const gradeCounts = { "A+": 0, "A": 0, "B": 0, "C": 0, "D": 0 };
  const gradeScores = { "A+": 100, "A": 90, "B": 80, "C": 70, "D": 60 };
  const validGrades = ["A+", "A", "B", "C", "D"];
  
  let totalScore = 0;
  const list = [];
  const categoryStats = {};
  const headers = rData.length > 0 ? rData[0] : [];
  const emailIndex = headers.indexOf("Trainee Email");
  
  // Trainer Stats Aggregation Object
  const trainerStats = {};
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  
  if (rData.length > 1) {
    for (let i = 1; i < rData.length; i++) {
      const r = rData[i];
      const grade = safeString(r[8]); 
      
      if (gradeCounts.hasOwnProperty(grade)) gradeCounts[grade]++;
      else gradeCounts["D"]++;
      
      totalScore += (gradeScores[grade] || 0);
      
      // --- START TRAINER AGGREGATION ---
      const tName = safeString(r[5]);
      if(tName) {
         if(!trainerStats[tName]) {
            trainerStats[tName] = {
               total: 0,
               branches: new Set(),
               firstHalf: 0,
               days: new Set(),
               monthlyTrainees: {} // TraineeID -> Set of Types
            };
         }
         
         const entry = trainerStats[tName];
         entry.total++; // Total Evaluations
         if(r[6]) entry.branches.add(safeString(r[6])); // Branch
         if(safeString(r[3]) === "1st Half") entry.firstHalf++; // Total 1st Half
         
         if(r[9]) {
            const d = new Date(r[9]);
            entry.days.add(d.toDateString()); // Unique Days
            
            // Monthly Completion Check
            if(d.getMonth() === currentMonth && d.getFullYear() === currentYear) {
               const tId = safeString(r[1]);
               if(!entry.monthlyTrainees[tId]) entry.monthlyTrainees[tId] = new Set();
               entry.monthlyTrainees[tId].add(safeString(r[3])); // Add Eval Type
            }
         }
      }
      // --- END TRAINER AGGREGATION ---
      
      let comments = [];
      let rowCatStats = {};

      for (let j = 11; j < r.length; j++) {
         const val = safeString(r[j]);
         const headerName = safeString(headers[j]); 
         
         if (!headerName || val === "") continue;

         if (headerName === "Trainee Email") continue;

         const cat = questionTextToCategory[headerName];
         const numericVal = gradeScores[val];

         if (cat && numericVal !== undefined) {
            if (!categoryStats[cat]) categoryStats[cat] = { sum: 0, count: 0 };
            categoryStats[cat].sum += numericVal;
            categoryStats[cat].count++;

            if (!rowCatStats[cat]) rowCatStats[cat] = { sum: 0, count: 0 };
            rowCatStats[cat].sum += numericVal;
            rowCatStats[cat].count++;
         }

         if (val && !validGrades.includes(val) && typeof val === 'string' && val.length > 2) {
             let cleanVal = val;
             if (headerName && cleanVal.startsWith(headerName + ":")) {
                 cleanVal = cleanVal.substring(headerName.length + 1).trim();
             } 
             else if (cleanVal.startsWith("Comment:")) {
                 cleanVal = cleanVal.substring(8).trim();
             }
             else if (cleanVal.startsWith("Remarks:")) {
                 cleanVal = cleanVal.substring(8).trim();
             }
             comments.push(cleanVal); 
         }
      }

      const rowCatScores = {};
      Object.keys(rowCatStats).forEach(c => {
         rowCatScores[c] = Math.round(rowCatStats[c].sum / rowCatStats[c].count);
      });

      const email = emailIndex > -1 ? safeString(r[emailIndex]) : "";

      list.push({
        id: safeString(r[0]), 
        traineeId: safeString(r[1]),
        traineeName: safeString(r[2]),
        evaluationType: safeString(r[3]),
        position: safeString(r[4]),
        trainerName: safeString(r[5]),
        branch: safeString(r[6]),
        overallGrade: safeString(r[8]),
        date: safeString(r[9]),
        pdfLink: safeString(r[10]),
        email: email,
        comments: comments.join("; "),
        categoryScores: rowCatScores
      });
    }
  }
  
  // Calculate Final Trainer Overview Array
  const trainerOverview = Object.keys(trainerStats).map(name => {
     const data = trainerStats[name];
     // Calculate completed monthly (both halves)
     let completedMonthly = 0;
     Object.values(data.monthlyTrainees).forEach(set => {
        if(set.has("1st Half") && set.has("2nd Half")) completedMonthly++;
     });
     
     return {
        name: name,
        totalTM: trainerTMs[name] || 0,
        totalEvaluations: data.total,
        totalBranches: data.branches.size,
        completedMonthly: completedMonthly,
        totalFirstHalf: data.firstHalf,
        totalDays: data.days.size
     };
  });
  
  const categoryPerformance = Object.keys(categoryStats).map(k => ({
    category: k,
    score: Math.round(categoryStats[k].sum / categoryStats[k].count) || 0
  }));
  
  if (categoryPerformance.length === 0) categoryPerformance.push({ category: "No Data", score: 0 });
  
  const distribution = Object.keys(gradeCounts).map(k => ({ name: k, value: gradeCounts[k] }));
  const count = list.length;
  const avg = count ? Math.round(totalScore / count) : 0;
  
  return {
    total: count,
    average: avg,
    distribution: distribution,
    categoryPerformance: categoryPerformance,
    list: list.reverse(),
    trainerOverview: trainerOverview // Return new stats
  };
}