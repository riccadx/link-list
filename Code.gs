/**
 * Google Apps Script Server Backend for Project & Systems Hub
 * Lab 305 Unified Directory
 */

// Define admin passcode (change this as needed)
const ADMIN_PASSCODE = "admin305";

// If you created this script directly in script.google.com (standalone),
// enter your Google Sheet's ID below (found in the sheet's browser URL).
// Example: const SPREADSHEET_ID = "1aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567890";
const SPREADSHEET_ID = "1HH_vr4eXt_1z7iiaEuAiXl4XI1EXgolTO-uBRoYhx80";

/**
 * Serve the web page
 */
function doGet(e) {
  // If accessed as an API, return JSON data
  if (e && e.parameter && e.parameter.api === 'true') {
    try {
      const data = getProjects();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Otherwise serve the HTML template (legacy/fallback mode)
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Project & Systems Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handle API POST requests from Vercel frontend
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    
    if (payload.action === 'save') {
      saveProject(payload.data);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } 
    
    if (payload.action === 'delete') {
      deleteProject(payload.internalId, payload.adminPasscode);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (payload.action === 'verify') {
      const isValid = verifyAdmin(payload.passcode);
      return ContentService.createTextOutput(JSON.stringify({ success: true, isValid: isValid }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (payload.action === 'recordClick') {
      recordClick(payload.internalId);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (payload.action === 'addComment') {
      addComment(payload.internalId, payload.commentData);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (payload.action === 'deleteComment') {
      deleteComment(payload.internalId, payload.commentTimestamp, payload.commentText, payload.adminPasscode);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    throw new Error("Unknown action: " + payload.action);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Increment click count for a project
 */
function recordClick(internalId) {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
  
  // Find or create Clicks column
  let clicksColIdx = headers.findIndex(h => h.toLowerCase() === "clicks");
  if (clicksColIdx === -1) {
    sheet.getRange(1, lastCol + 1).setValue("Clicks").setFontWeight("bold").setBackground("#e2e8f0");
    clicksColIdx = lastCol; // 0-indexed index of new column
  }
  
  const idColIdx = headers.findIndex(h => h.toLowerCase() === "_internal_id" || h === "ID");
  if (idColIdx === -1) return; // Cannot map
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  // Find row with matching internalId
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][idColIdx]).trim();
    if (rowId === String(internalId).trim()) {
       // Increment clicks
       const currentClicks = parseInt(data[i][clicksColIdx]) || 0;
       sheet.getRange(i + 1, clicksColIdx + 1).setValue(currentClicks + 1);
       break;
    }
  }
}

/**
 * Include other HTML/CSS/JS files inside the template
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get sheet object, initializing headers if empty
 */
function getSheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    // Ignore error and fall back to ID
  }
  
  if (!ss && SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      throw new Error("Could not open spreadsheet by ID. Make sure SPREADSHEET_ID is correct and this script has permissions: " + e.toString());
    }
  }
  
  if (!ss) {
    throw new Error(
      "No active spreadsheet found.\n\n" +
      "HOW TO FIX THIS:\n" +
      "1. Either create this Apps Script inside a Google Sheet (Extensions > Apps Script),\n" +
      "2. Or create a Google Sheet, copy its ID from the browser URL, and paste it at the top of Code.gs as: const SPREADSHEET_ID = 'your_sheet_id_here';"
    );
  }

  let sheet = ss.getSheetByName("Projects");
  if (!sheet) {
    sheet = ss.insertSheet("Projects");
  }
  
  // Set headers if sheet is brand new / empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Title", "URL", "Description", "Submitter", "Timestamp"]);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  
  return sheet;
}

/**
 * Fetch all projects dynamically reading headers
 */
function getProjects() {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Get exact headers even if sheet is empty of data
    let headers = [];
    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
    }
    
    if (lastRow <= 1) {
      return { headers: headers.filter(h => h), data: [] };
    }
    
    // Find an ID column if exists
    const idIdx = headers.findIndex(h => h.toLowerCase() === "id" || h.toLowerCase() === "identifier" || h.toLowerCase() === "no.");
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    
    const data = values.map((row, rowIndex) => {
      let obj = {};
      // Provide an internal tracking ID even if the user has no ID column
      obj._internal_id = (idIdx !== -1 && row[idIdx]) ? row[idIdx] : "ROW-" + (rowIndex + 2);
      
      headers.forEach((h, i) => {
        if (h) {
          let val = row[i];
          if (val instanceof Date) {
            val = val.toISOString(); // Fix Apps Script serialization crash
          }
          obj[h] = val !== undefined ? val : "";
        }
      });
      return obj;
    });
    
    return { headers: headers.filter(h => h), data: data };
  } catch (error) {
    throw new Error("Failed to load systems: " + error.toString());
  }
}

/**
 * Run this function ONCE from the editor to force the Authorization popup for Drive and Email (MailApp)
 */
function authorizeServices() {
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  console.log("Drive & Mail Authorization Successful! Remaining daily email quota: " + MailApp.getRemainingDailyQuota());
}

/**
 * Helper to upload base64 image to Google Drive and return public URL
 */
function uploadBase64ToDrive(base64Data, filename) {
  try {
    const splitIndex = base64Data.indexOf("base64,");
    if (splitIndex !== -1) {
       const typePart = base64Data.substring(5, splitIndex - 1);
       const base64Str = base64Data.substring(splitIndex + 7);
       
       const blob = Utilities.newBlob(Utilities.base64Decode(base64Str), typePart, filename);
       
       let folders = DriveApp.getFoldersByName("Lab 305 Logos");
       let folder;
       if (folders.hasNext()) {
         folder = folders.next();
       } else {
         folder = DriveApp.createFolder("Lab 305 Logos");
         try {
           folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         } catch(err) { console.log("Folder sharing restricted by admin", err); }
       }
       
       const file = folder.createFile(blob);
       try {
         file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       } catch(err) { console.log("File sharing restricted by admin", err); }
       
       return "https://drive.google.com/uc?export=view&id=" + file.getId();
    }
  } catch(e) {
    // Drive failed (likely org permissions), fallback to raw base64
    return base64Data;
  }
  return base64Data;
}

/**
 * Save or update a project row
 */
function saveProject(project) {
  try {
    // Intercept base64 images and upload to Drive
    Object.keys(project).forEach(k => {
       const val = project[k];
       if (val && typeof val === 'string' && val.startsWith("data:image/")) {
           project[k] = uploadBase64ToDrive(val, "Upload_" + Date.now());
       }
    });

    const sheet = getSheet();
    let lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
       sheet.appendRow(["ID"]);
       lastCol = 1;
    }
    const lastRow = sheet.getLastRow();
    
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
    
    // Dynamically append any missing headers (e.g. from new Edit features)
    const incomingKeys = Object.keys(project).filter(k => k !== '_internal_id' && k !== '');
    let hasNewHeaders = false;
    incomingKeys.forEach(k => {
       if (!headers.includes(k) && !headers.some(h => h.toLowerCase() === k.toLowerCase())) {
          headers.push(k);
          hasNewHeaders = true;
       }
    });
    
    if (hasNewHeaders) {
       sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
       sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
       lastCol = headers.length;
    }
    
    const idIdx = headers.findIndex(h => h.toLowerCase() === "id" || h.toLowerCase() === "identifier" || h.toLowerCase() === "no.");
    
    let targetRow = -1;
    let internalId = project._internal_id;
    
    // Find row by internal ID
    if (internalId && String(internalId).startsWith("ROW-")) {
       targetRow = parseInt(String(internalId).split("-")[1], 10);
    } else if (internalId && idIdx !== -1 && lastRow > 1) {
       const idValues = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
       for (let i = 0; i < idValues.length; i++) {
         if (String(idValues[i][0]) === String(internalId)) {
           targetRow = i + 2;
           break;
         }
       }
    }
    
    if (targetRow !== -1) {
      // Edit existing row
      headers.forEach((h, i) => {
        // If the header exists and we have data for it, update it (or clear it)
        if (h && typeof project[h] !== 'undefined') {
           let val = project[h];
           
           // Special handling for the rolling edit log
           if (h.toLowerCase() === "edit log") {
              const existingVal = sheet.getRange(targetRow, i + 1).getValue();
              const dateStr = new Date().toLocaleString();
              const who = project["Last Edited By"] || "Unknown";
              val = `[${dateStr}] ${who}: ${val}`;
              if (existingVal) {
                 val = val + "\n\n" + existingVal; // append to history
              }
           }
           
           sheet.getRange(targetRow, i + 1).setValue(val);
        }
      });
    } else {
      // Add new row matching headers structure
      internalId = "ID-" + Utilities.getUuid().substring(0, 8);
      const newRow = new Array(headers.length).fill("");
      
      headers.forEach((h, i) => {
         if (idIdx === i) {
            newRow[i] = internalId;
         } else if (h && typeof project[h] !== 'undefined') {
            newRow[i] = project[h];
         }
         
         // Auto-timestamp if column is named Timestamp
         if (h && h.toLowerCase() === "timestamp" && !project[h]) {
            newRow[i] = new Date();
         }
      });
      
      sheet.appendRow(newRow);
    }
    return { success: true };
  } catch (error) {
    throw new Error("Failed to save system: " + error.toString());
  }
}

/**
 * Delete a project row by ID dynamically
 */
function deleteProject(internalId, passcode) {
  if (passcode !== ADMIN_PASSCODE) {
    throw new Error("Unauthorized action.");
  }
  
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, reason: "Empty sheet" };
    
    if (String(internalId).startsWith("ROW-")) {
      const targetRow = parseInt(String(internalId).split("-")[1], 10);
      sheet.deleteRow(targetRow);
      return { success: true };
    }
    
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
    const idIdx = headers.findIndex(h => h.toLowerCase() === "id" || h.toLowerCase() === "identifier" || h.toLowerCase() === "no.");
    
    if (idIdx === -1) throw new Error("ID column not found in sheet headers.");
    
    const dataRange = sheet.getRange(2, idIdx + 1, lastRow - 1, 1);
    const values = dataRange.getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === String(internalId)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, reason: "ID not found" };
  } catch (error) {
    throw new Error("Failed to delete system: " + error.toString());
  }
}

/**
 * Validate admin passcode
 */
function verifyAdmin(passcode) {
  return passcode === ADMIN_PASSCODE;
}

/**
 * Add a comment to a project and notify the developer
 */
function addComment(internalId, commentData) {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
  
  // Ensure Comments column exists
  let commentsColIdx = headers.findIndex(h => h.toLowerCase() === "comments" || h === "コメント");
  if (commentsColIdx === -1) {
    sheet.getRange(1, lastCol + 1).setValue("Comments").setFontWeight("bold").setBackground("#e2e8f0");
    commentsColIdx = lastCol;
    headers.push("Comments");
  }
  
  const idColIdx = headers.findIndex(h => h.toLowerCase() === "_internal_id" || h.toLowerCase() === "id" || h.toLowerCase() === "identifier");
  if (idColIdx === -1) throw new Error("Cannot find ID column");
  
  const titleColIdx = headers.findIndex(h => {
    const l = h.toLowerCase();
    return l === "title" || l === "タイトル" || l === "system name" || l === "システム名" || l === "name" || l === "名前";
  });

  // Look for developer email column
  const devEmailColIdx = headers.findIndex(h => {
    const l = h.toLowerCase();
    return l.includes("email") || l.includes("メール") || l.includes("contact") || l.includes("developer") || 
           l.includes("author") || l.includes("submitter") || l.includes("owner") || 
           l.includes("作成者") || l.includes("担当者") || l.includes("連絡先");
  });
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][idColIdx]).trim();
    if (rowId === String(internalId).trim()) {
       // Parse existing comments
       let comments = [];
       try {
         const cellData = data[i][commentsColIdx];
         if (cellData) comments = JSON.parse(cellData);
       } catch (e) {
         comments = [];
       }
       
       // Filter out comments older than 30 days
       const thirtyDaysAgo = Date.now() - (30 * 24 * 60 * 60 * 1000);
       comments = comments.filter(c => c && c.timestamp >= thirtyDaysAgo);
       
       // Append new comment
       const newComment = {
         name: commentData.name || "Anonymous",
         email: commentData.email || "",
         text: commentData.text,
         timestamp: Date.now()
       };
       comments.push(newComment);
       
       // Save back to sheet
       sheet.getRange(i + 1, commentsColIdx + 1).setValue(JSON.stringify(comments));
       
       // Determine developer email
       let devEmail = "";
       if (devEmailColIdx !== -1 && data[i][devEmailColIdx]) {
         const candidate = String(data[i][devEmailColIdx]).trim();
         if (candidate.includes("@")) {
           devEmail = candidate;
         }
       }
       
       // Fallback: search all columns in this row for any valid email address
       if (!devEmail) {
         const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
         for (let col = 0; col < data[i].length; col++) {
           if (col === commentsColIdx) continue;
           const cellVal = String(data[i][col]).trim();
           const match = cellVal.match(emailRegex);
           if (match) {
             devEmail = match[0];
             break;
           }
         }
       }

       // Send email notification to developer if email exists
       const projTitle = (titleColIdx !== -1 && data[i][titleColIdx]) ? data[i][titleColIdx] : "a project";
       
       if (devEmail && devEmail.includes("@")) {
         const subject = `New feedback on your project: ${projTitle}`;
         const body = `Hello,\n\nSomeone just left new feedback on your project (${projTitle}) in the Lab 305 Directory!\n\n`
                    + `From: ${newComment.name}\n`
                    + `Comment: "${newComment.text}"\n\n`
                    + (newComment.email ? `User Email for Reply: ${newComment.email}\n(You can reply directly to this email to answer them!)\n\n` : `\n\n`)
                    + `Best regards,\nLab 305 Project Hub`;
                    
         try {
           const emailOptions = {
             to: devEmail,
             subject: subject,
             body: body
           };
           if (newComment.email && newComment.email.includes("@")) {
             emailOptions.replyTo = newComment.email;
           }
           MailApp.sendEmail(emailOptions);
           console.log("Notification email sent successfully to: " + devEmail);
         } catch(emailErr) {
           console.error("Failed to send notification email to " + devEmail + ": " + emailErr.toString());
         }
       } else {
         console.warn("No valid developer email address found for project: " + projTitle);
       }
       break;
    }
  }
}

/**
 * Delete a specific comment from a project (Admin only)
 */
function deleteComment(internalId, timestamp, text, passcode) {
  if (!verifyAdmin(passcode)) {
    throw new Error("Unauthorized action. Invalid admin passcode.");
  }
  
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
  
  let commentsColIdx = headers.findIndex(h => h.toLowerCase() === "comments" || h === "コメント");
  if (commentsColIdx === -1) return { success: false, error: "No comments column found" };
  
  const idColIdx = headers.findIndex(h => h.toLowerCase() === "_internal_id" || h.toLowerCase() === "id" || h.toLowerCase() === "identifier");
  if (idColIdx === -1) throw new Error("Cannot find ID column");
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  const thirtyDaysAgo = Date.now() - (30 * 24 * 60 * 60 * 1000);
  
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][idColIdx]).trim();
    if (rowId === String(internalId).trim()) {
       let comments = [];
       try {
         const cellData = data[i][commentsColIdx];
         if (cellData) comments = JSON.parse(cellData);
       } catch (e) {
         comments = [];
       }
       
       comments = comments.filter(c => {
         if (!c) return false;
         // Auto remove if older than 30 days
         if (c.timestamp && c.timestamp < thirtyDaysAgo) return false;
         
         // Match by timestamp if present
         if (timestamp && c.timestamp && String(c.timestamp) === String(timestamp)) return false;
         // Match by text if timestamp is missing or ambiguous
         if (text && (c.text === text || (typeof c === 'string' && c === text))) return false;
         
         return true;
       });
       
       sheet.getRange(i + 1, commentsColIdx + 1).setValue(JSON.stringify(comments));
       return { success: true };
    }
  }
  return { success: false, error: "Project not found" };
}
