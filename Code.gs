/**
 * Google Apps Script Server Backend for Project & Systems Hub
 * Lab 305 Unified Directory
 */

// Bound Google Sheet ID
const SPREADSHEET_ID = "1HH_vr4eXt_1z7iiaEuAiXl4XI1EXgolTO-uBRoYhx80";

// Salt used for one-way SHA-256 hash verification
const PASSCODE_SALT = "Lab305_Unified_Directory_Salt";

// One-way cryptographic hash fallback (Plaintext password is NEVER stored in code or sheets)
const DEFAULT_PASSCODE_HASH = "DQothiIWFjLRjKb9ZZ7l4QwQ9KOSu+8koUDqiAXDz/E=";

/**
 * Hash a passcode using SHA-256 with project salt
 */
function hashPasscode(passcode) {
  if (!passcode) return "";
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(passcode).trim() + PASSCODE_SALT,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(digest);
}

/**
 * Retrieve the active admin passcode hash from Script Properties or fallback
 */
function getAdminPasscodeHash() {
  try {
    const hash = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSCODE_HASH");
    if (hash && hash.trim()) return hash.trim();
    
    const plain = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSCODE");
    if (plain && plain.trim()) return hashPasscode(plain.trim());
  } catch(e) {}
  return DEFAULT_PASSCODE_HASH;
}

/**
 * Admin utility: Update admin passcode securely in Script Properties.
 * Run this function from the Apps Script editor to change your password without ever touching Code.gs or Google Sheets.
 * Example: setAdminPasscode("myNewSecurePassword123");
 */
function setAdminPasscode(newPasscode) {
  if (!newPasscode || typeof newPasscode !== 'string') {
    throw new Error("Please specify a valid string passcode: setAdminPasscode('your_new_password')");
  }
  const hash = hashPasscode(newPasscode);
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSCODE_HASH", hash);
  // Remove any legacy plaintext property
  try { PropertiesService.getScriptProperties().deleteProperty("ADMIN_PASSCODE"); } catch(e) {}
  console.log("Admin passcode successfully hashed and stored in Google Cloud Script Properties!");
  return true;
}

/**
 * Admin utility: Sanitize Google Sheet columns.
 * Permanently removes any column from the Google Sheet whose header contains "password", "passcode", "token", "secret", or "credential".
 * Run this function from the Apps Script editor to ensure your Google Sheet never displays or holds admin credentials.
 */
function sanitizeSheetColumns() {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    console.log("Sheet has no columns.");
    return "Sheet is empty.";
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
  const sensitivePatterns = ["passcode", "password", "token", "secret", "credential"];
  let removedCount = 0;
  for (let i = headers.length - 1; i >= 0; i--) {
    const h = headers[i].toLowerCase();
    if (sensitivePatterns.some(p => h.includes(p))) {
      console.log("Purging sensitive column from Google Sheet: " + headers[i]);
      sheet.deleteColumn(i + 1);
      removedCount++;
    }
  }
  console.log("Sanitization complete. Purged " + removedCount + " sensitive column(s).");
  return "Purged " + removedCount + " sensitive column(s) from Google Sheet.";
}

/**
 * Generate a cryptographically signed HMAC session token
 */
function generateAdminToken() {
  const timestamp = Date.now().toString();
  const nonce = Utilities.getUuid().substring(0, 8);
  const data = timestamp + ":" + nonce;
  const secret = getAdminPasscodeHash();
  const rawSig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, data, secret);
  const signature = Utilities.base64EncodeWebSafe(rawSig);
  const token = data + ":" + signature;
  
  try {
    CacheService.getScriptCache().put("SESSION_" + token, "valid", 14400); // 4 hours
  } catch(e) {}
  
  return token;
}

/**
 * Verify HMAC session token
 */
function verifyAdminToken(token) {
  if (!token || typeof token !== 'string') return false;
  
  // Check revocation cache
  try {
    const isRevoked = CacheService.getScriptCache().get("REVOKED_" + token);
    if (isRevoked === "true") return false;
  } catch(e) {}
  
  // Fast path: cached valid session
  try {
    const cached = CacheService.getScriptCache().get("SESSION_" + token);
    if (cached === "valid") return true;
  } catch(e) {}
  
  // Signature verification
  const parts = token.split(":");
  if (parts.length !== 3) return false;
  
  const timestamp = parseInt(parts[0], 10);
  if (isNaN(timestamp)) return false;
  
  // 4 hour expiration window
  const fourHours = 4 * 60 * 60 * 1000;
  if (Date.now() - timestamp > fourHours || timestamp > Date.now() + 60000) {
    return false;
  }
  
  const data = parts[0] + ":" + parts[1];
  const secret = getAdminPasscodeHash();
  const rawSig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, data, secret);
  const expectedSig = Utilities.base64EncodeWebSafe(rawSig);
  
  if (parts[2] === expectedSig) {
    try {
      CacheService.getScriptCache().put("SESSION_" + token, "valid", 14400);
    } catch(e) {}
    return true;
  }
  
  return false;
}

/**
 * Revoke an admin session token on logout
 */
function revokeAdminToken(token) {
  if (!token) return;
  try {
    CacheService.getScriptCache().remove("SESSION_" + token);
    CacheService.getScriptCache().put("REVOKED_" + token, "true", 14400);
  } catch(e) {}
}

/**
 * Validate either a signed session token or legacy passcode
 */
function verifyAdminOrToken(credential) {
  if (!credential) return false;
  if (verifyAdminToken(credential)) return true;
  return verifyAdmin(credential);
}

/**
 * Serve the web page
 */
function doGet(e) {
  // If accessed as an API, return JSON data (filters protected columns unless token provided)
  if (e && e.parameter && e.parameter.api === 'true') {
    try {
      const token = e.parameter.token || e.parameter.authToken;
      const data = getProjects(token);
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
    const token = payload.authToken || payload.adminToken || payload.adminPasscode;
    
    // Login / Authentication handshake: produces signed token, NEVER persists password
    if (payload.action === 'login' || payload.action === 'verify') {
      const isValid = verifyAdmin(payload.passcode);
      if (isValid) {
        const sessionToken = generateAdminToken();
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          isValid: true,
          token: sessionToken,
          expiresIn: 14400
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          isValid: false,
          error: "Invalid administrator credentials."
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (payload.action === 'logout') {
      revokeAdminToken(token);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.action === 'getProjects') {
      const data = getProjects(token);
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (payload.action === 'save') {
      saveProject(payload.data, token);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } 
    
    if (payload.action === 'delete') {
      deleteProject(payload.internalId, token);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
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
      deleteComment(payload.internalId, payload.commentTimestamp, payload.commentText, token);
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
function getProjects(authToken) {
  try {
    const isAdmin = verifyAdminOrToken(authToken);
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Get exact headers even if sheet is empty of data
    let headers = [];
    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
    }
    
    // Absolute ban on exposing any passwords, secrets, or auth tokens to ANY client
    const isSensitiveHeader = (h) => {
      const low = h.toLowerCase();
      return low.includes("passcode") || low.includes("password") || low.includes("token") || low.includes("secret") || low.includes("credential");
    };

    // Filter protected internal headers (edit log, last edited by, admin notes) from public view unless admin token is validated
    const isProtectedHeader = (h) => {
      const low = h.toLowerCase();
      return low === "edit log" || low === "last edited by" || low.includes("admin notes") || isSensitiveHeader(h);
    };

    const activeHeaders = headers.filter(h => h && !isSensitiveHeader(h) && (isAdmin || !isProtectedHeader(h)));
    
    if (lastRow <= 1) {
      return { headers: activeHeaders, data: [], isAdmin: isAdmin };
    }
    
    // Find an ID column if exists
    const idIdx = headers.findIndex(h => h.toLowerCase() === "id" || h.toLowerCase() === "identifier" || h.toLowerCase() === "no.");
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    
    const data = values.map((row, rowIndex) => {
      let obj = {};
      // Provide an internal tracking ID even if the user has no ID column
      obj._internal_id = (idIdx !== -1 && row[idIdx]) ? row[idIdx] : "ROW-" + (rowIndex + 2);
      
      activeHeaders.forEach((h) => {
        const colIdx = headers.indexOf(h);
        if (colIdx !== -1) {
          let val = row[colIdx];
          if (val instanceof Date) {
            val = val.toISOString(); // Fix Apps Script serialization crash
          }
          obj[h] = val !== undefined ? val : "";
        }
      });
      return obj;
    });
    
    return { headers: activeHeaders, data: data, isAdmin: isAdmin };
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
function saveProject(project, tokenOrPasscode) {
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
    
    // Security: Filter out internal and sensitive keys so they NEVER become headers or cells in Google Sheet
    const SENSITIVE_KEY_PATTERNS = ["_internal_id", "token", "passcode", "password", "secret", "credential", "auth"];
    const isSensitiveField = (k) => {
      if (!k) return true;
      const lower = k.toString().toLowerCase().trim();
      return SENSITIVE_KEY_PATTERNS.some(p => lower.includes(p));
    };

    // Dynamically append any missing headers (excluding sensitive fields)
    const incomingKeys = Object.keys(project).filter(k => k && k !== '_internal_id' && !isSensitiveField(k));
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
      // Modifying an existing row: MUST BE AUTHENTICATED
      if (!verifyAdminOrToken(tokenOrPasscode)) {
        throw new Error("Unauthorized action. Valid administrator session required to modify existing projects.");
      }

      // Edit existing row
      headers.forEach((h, i) => {
        // If the header exists, is not sensitive, and we have data for it, update it
        if (h && !isSensitiveField(h) && typeof project[h] !== 'undefined') {
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
         } else if (h && !isSensitiveField(h) && typeof project[h] !== 'undefined') {
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
function deleteProject(internalId, tokenOrPasscode) {
  if (!verifyAdminOrToken(tokenOrPasscode)) {
    throw new Error("Unauthorized action. Valid administrator session required to delete projects.");
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
  if (!passcode) return false;
  return hashPasscode(passcode) === getAdminPasscodeHash();
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
       
       // Save back to sheet (website comments array stays text-only)
       sheet.getRange(i + 1, commentsColIdx + 1).setValue(JSON.stringify(comments));
       
       // Process attached images for email notification (up to 3 images)
       let attachments = [];
       const rawImages = (commentData && commentData.images) ? commentData.images : (commentData && commentData.image ? [commentData.image] : []);
       
       rawImages.slice(0, 3).forEach((img, idx) => {
         if (img && img.data) {
           try {
             const splitIdx = img.data.indexOf("base64,");
             if (splitIdx !== -1) {
               const mimeType = img.type || "image/png";
               const base64Str = img.data.substring(splitIdx + 7);
               const blob = Utilities.newBlob(Utilities.base64Decode(base64Str), mimeType, img.filename || `attached_photo_${idx + 1}.png`);
               attachments.push(blob);
             }
           } catch(attachErr) {
             console.error("Error creating image attachment: " + attachErr.toString());
           }
         }
       });
       
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
         const subject = `【Lab 305】プロジェクト「${projTitle}」に新しいフィードバックが届きました`;
          
         let body = `こんにちは！日頃の開発・運用お疲れ様です。\n`
                  + `(Hello, thank you for your hard work!)\n\n`
                  + `Lab 305 プロジェクトディレクトリに掲載中のあなたのプロジェクト「${projTitle}」へ、新しいフィードバック・コメントが届きました！\n\n`
                  + `--------------------------------------------------\n`
                  + `■ 投稿者 (From): ${newComment.name}\n`
                  + `■ コメント内容 (Comment):\n"${newComment.text}"\n`
                  + `--------------------------------------------------\n\n`;
                  
         if (attachments.length > 0) {
           body += `📎 添付画像（${attachments.length}件）が含まれています。本メールの添付ファイルをご確認ください。\n`
                 + `(${attachments.length} photo attachment(s) included with this email!)\n\n`;
         }
         
         if (newComment.email) {
           body += `■ 返信用メールアドレス (Reply-to Email): ${newComment.email}\n`
                 + `（※このメールに直接返信すると、投稿者にそのまま回答を送ることができます！）\n`
                 + `(You can reply directly to this email to answer them!)\n\n`;
         }
         
         body += `引き続き開発・管理よろしくお願いいたします。\n\n`
               + `敬具\n`
               + `Lab 305 Project Hub`;
                    
         try {
           const emailOptions = {
             to: devEmail,
             subject: subject,
             body: body
           };
           if (newComment.email && newComment.email.includes("@")) {
             emailOptions.replyTo = newComment.email;
           }
           if (attachments.length > 0) {
             emailOptions.attachments = attachments;
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
function deleteComment(internalId, timestamp, text, tokenOrPasscode) {
  if (!verifyAdminOrToken(tokenOrPasscode)) {
    throw new Error("Unauthorized action. Valid administrator session required to delete comments.");
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
