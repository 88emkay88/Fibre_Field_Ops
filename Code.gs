/**
 * ============================================================================
 * FibreGems Field Operations Center — Google Apps Script Backend (Code.gs)
 * ============================================================================
 * 
 * INSTRUCTIONS FOR DEPLOYMENT:
 * 1. Open your Google Sheet -> Extensions -> Apps Script
 * 2. Replace the entire contents of Code.gs with this code.
 * 3. In the function dropdown, select `migrateAndExtendSheets` and click Run once.
 * 4. Click "Deploy" -> "Manage deployments" -> Edit (pencil icon) -> "New version" -> Deploy.
 *    (Ensure "Execute as: Me" and "Who has access: Anyone").
 * ============================================================================
 */
 

// ════════════════════════════════════════════════════════════════════════════
// 1. GLOBAL CONSTANTS
// ════════════════════════════════════════════════════════════════════════════
const FOLDER_NAME = "FibreGems_Uploads";
const TIMEZONE = "Africa/Johannesburg"; // SAST (UTC+2)
const LOCATIONIQ_API_KEY = "pk.1be07ee2080691339d8fc4f1712dbc95";

// ════════════════════════════════════════════════════════════════════════════
// 2. HTTP GET HANDLER
// ════════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : null;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "getSuperAdminData" || action === "getPortalData" || action === "getLeaderData" || action === "getAgentData") {
      return getSuperAdminData(ss);
    }
    if (action === "getObjectivesData") {
      return getObjectivesDataHandler(ss);
    }
    return createJsonResponse({ status: "success", message: "FibreGems Operations API is active." });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 3. HTTP POST HANDLER
// ════════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "No post data received." });
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    switch (action) {
      // ── Authentication ──
      case "loginUser":
        return createJsonResponse(loginUser(ss, payload.email, payload.password));
      case "registerUser":
        return createJsonResponse(registerUser(ss, payload.email, payload.firstName, payload.lastName, payload.password));
      case "addTeamLeader":
        return createJsonResponse(addTeamLeader(ss, payload.firstName, payload.lastName, payload.email, payload.password, payload.accountType));
      case "deleteTeamLeader":
        return createJsonResponse(deleteTeamLeader(ss, payload.email));

      // ── Field Check-Ins ──
      case "submitAgentSignOn":
      case "submitAgentClockIn":
        return handleAgentSignOn(ss, payload);
      case "submitLeaderCheckIn":
        return handleLeaderCheckIn(ss, payload);

      // ── Objectives ──
      case "setWeeklyObjective":
        return handleSetWeeklyObjective(ss, payload);
      case "requestObjective":
        return handleRequestObjective(ss, payload);
      case "approveObjective":
        return handleApproveObjective(ss, payload);

      // ── Agent Performance & Strikes ──
      case "logAgentStats":
      case "logAgentIssue":
        return handleLogAgentStats(ss, payload);

      // ── Reverse Geocoding ──
      case "getReadableLocation":
        const locationName = getReadableLocation(payload.lat, payload.lon);
        return createJsonResponse({
          status: "success",
          display_name: locationName,
          address: locationName
        });

      default:
        return createJsonResponse({ status: "error", message: "Invalid action: " + action });
    }
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════════
// 4. SCHEMA MIGRATION & EXTENSION (Run once)
// ════════════════════════════════════════════════════════════════════════════
/**
 * Run this function ONCE from the Apps Script editor to safely
 * extend your existing sheets with new columns without losing any data.
 */
function migrateAndExtendSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return Logger.log("❌ No active spreadsheet found.");

  // 1. Objectives Sheet (adds 'Target Houses' at column 10)
  const objSheetName = ss.getSheetByName("Objectives") ? "Objectives" : "WeeklyObjectives";
  extendSheetHeaders(ss, objSheetName, [
    "Date", "View Type", "Target Date", "Leader Name",
    "Assigned Location", "Focus Areas", "Target Metrics", "Status", "Requested By",
    "Target Houses"
  ]);

  // 2. AgentSignOns Sheet
  extendSheetHeaders(ss, "AgentSignOns", [
    "Date", "Time", "Agent Name", "Leader Name", "Location",
    "Photo URL", "Is Late (After 10)", "GPS", "Location Detail", "Notes"
  ]);

  // 3. AgentStats Sheet
  extendSheetHeaders(ss, "AgentStats", [
    "Date", "Agent Name", "Engagements", "Real Leads", "Payments",
    "Consecutive Misses (Strikes)", "Time", "Leader Name", "Issue Type", "Notes"
  ]);

  // 4. LeaderCheckIns Sheet
  extendSheetHeaders(ss, "LeaderCheckIns", [
    "Date", "Time", "Leader Name", "Region", "Odometer/Vehicle Info", "GPS", "Selfie/Photo URL", "Location"
  ]);

  // 5. Users Sheet
  extendSheetHeaders(ss, "Users", [
    "Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Reset Token"
  ]);

  clearCache();
  Logger.log("✅ All sheets checked and successfully extended!");
}

function extendSheetHeaders(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(expectedHeaders);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight("bold").setBackground("#e2e8f0");
    Logger.log(`Created new sheet: ${sheetName}`);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  const missingHeaders = [];
  expectedHeaders.forEach(h => {
    if (!currentHeaders.includes(h)) {
      missingHeaders.push(h);
    }
  });

  if (missingHeaders.length > 0) {
    const startCol = lastCol + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]).setFontWeight("bold").setBackground("#e2e8f0");
    Logger.log(`Extended ${sheetName} with columns: ${missingHeaders.join(", ")}`);
  } else {
    Logger.log(`${sheetName} already has all required columns.`);
  }
}

/**
 * Initialises default Super Admin account if not existing.
 */
function setupLoginTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateUsersSheet(ss);
  const defaultEmail = "admin@fibregems.co.za";
  const defaultPassword = "AdminPassword123!";
  const salt = generateSalt(16);
  const hash = computeSHA256(defaultPassword, salt);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === defaultEmail) return Logger.log("Admin exists.");
  }
  sheet.appendRow([defaultEmail, "Super", "Admin", hash, salt, "Admin", "Verified", ""]);
  Logger.log("Default admin created: " + defaultEmail);
}

// ════════════════════════════════════════════════════════════════════════════
// 5. REVERSE GEOCODING (LocationIQ)
// ════════════════════════════════════════════════════════════════════════════
function getReadableLocation(lat, lon) {
  try {
    const url = `https://us1.locationiq.com/v1/reverse?key=${LOCATIONIQ_API_KEY}&lat=${lat}&lon=${lon}&format=json`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data && data.display_name) {
        return data.display_name;
      }
      return "ERROR: Missing display_name. Payload: " + response.getContentText();
    }
    return `API Error ${response.getResponseCode()}: ${response.getContentText()}`;
  } catch (err) {
    return `FETCH ERROR: ${err.toString()}`;
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 6. DRIVE & SHEETS UTILITIES
// ════════════════════════════════════════════════════════════════════════════
function getUploadFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

function saveBase64Image(base64Data, filenamePrefix) {
  if (!base64Data || !base64Data.startsWith("data:image")) return "";
  try {
    const parts = base64Data.split(",");
    const raw = parts[1];
    const mimeMatch = parts[0].match(/data:(.*?);base64/);
    const contentType = mimeMatch ? mimeMatch[1] : "image/jpeg";
    const ext = contentType.includes("png") ? "png" : "jpg";
    const decoded = Utilities.base64Decode(raw);
    const blob = Utilities.newBlob(decoded, contentType, `${filenamePrefix}_${new Date().getTime()}.${ext}`);
    const file = getUploadFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    return "";
  }
}

function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getOrCreateUsersSheet(ss) {
  return getOrCreateSheet(ss, "Users", [
    "Email", "First Name", "Last Name", "Password Hash", "Salt",
    "Account Type", "Status", "Reset Token"
  ]);
}

// ════════════════════════════════════════════════════════════════════════════
// 7. AUTHENTICATION & TEAM LEADER MANAGEMENT
// ════════════════════════════════════════════════════════════════════════════
function generateSalt(length) {
  const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let salt = "";
  for (let i = 0; i < length; i++) salt += charset.charAt(Math.floor(Math.random() * charset.length));
  return salt;
}

function computeSHA256(password, salt) {
  const rawData = password + salt;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawData, Utilities.Charset.UTF_8);
  let hexString = "";
  for (let i = 0; i < digest.length; i++) {
    let byte = digest[i];
    if (byte < 0) byte += 256;
    let hex = byte.toString(16);
    if (hex.length === 1) hex = "0" + hex;
    hexString += hex;
  }
  return hexString;
}

function registerUser(ss, email, firstName, lastName, password) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === emailLower) {
        return { status: "error", message: "Account already exists." };
      }
    }
    const salt = generateSalt(16);
    const hash = computeSHA256(password, salt);
    sheet.appendRow([emailLower, firstName, lastName, hash, salt, "Agent", "Verified", ""]);
    clearCache();
    return { status: "success", message: "Registration successful. Please log in." };
  } catch (err) {
    return { status: "error", message: "Server error: " + err.toString() };
  }
}

function loginUser(ss, email, password) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[0].toString().toLowerCase() === emailLower) {
        if (row[6] === "Blocked") return { status: "error", message: "Account blocked." };
        if (computeSHA256(password, row[4]) === row[3]) {
          return { status: "success", user: { email: row[0], firstName: row[1], lastName: row[2], accountType: row[5] } };
        } else {
          return { status: "error", message: "Invalid credentials." };
        }
      }
    }
    return { status: "error", message: "Invalid credentials." };
  } catch (err) {
    return { status: "error", message: "Server error: " + err.toString() };
  }
}

function addTeamLeader(ss, firstName, lastName, email, password, accountType) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === emailLower) {
        return { status: "error", message: "User with this email already exists." };
      }
    }

    if (!firstName || !lastName || !email || !password) {
      return { status: "error", message: "All fields are required." };
    }
    if (!email.includes("@")) {
      return { status: "error", message: "Invalid email address." };
    }
    if (password.length < 6) {
      return { status: "error", message: "Password must be at least 6 characters." };
    }

    const salt = generateSalt(16);
    const hash = computeSHA256(password, salt);
    
    sheet.appendRow([
      emailLower, 
      firstName, 
      lastName, 
      hash, 
      salt, 
      accountType || "Team Leader", 
      "Verified", 
      ""
    ]);
    
    clearCache();
    return { 
      status: "success", 
      message: `Team Leader ${firstName} ${lastName} added successfully.` 
    };
  } catch (err) {
    return { status: "error", message: "Server error: " + err.toString() };
  }
}

function deleteTeamLeader(ss, email) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();

    if (!emailLower) {
      return { status: "error", message: "Email is required." };
    }

    let deleted = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === emailLower) {
        sheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }

    if (!deleted) {
      return { status: "error", message: "Team Leader not found." };
    }

    clearCache();
    return { status: "success", message: `Team Leader ${email} removed successfully.` };
  } catch (err) {
    return { status: "error", message: "Server error: " + err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 8. OPERATIONAL HANDLERS
// ════════════════════════════════════════════════════════════════════════════
function handleAgentSignOn(ss, payload) {
  const sheet = getOrCreateSheet(ss, "AgentSignOns", [
    "Date", "Time", "Agent Name", "Leader Name", "Location",
    "Photo URL", "Is Late (After 10)", "GPS", "Location Detail", "Notes"
  ]);
  const now = new Date();
  const dateStr = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, TIMEZONE, "HH:mm:ss");
  const hours = parseInt(Utilities.formatDate(now, TIMEZONE, "HH"), 10);
  const minutes = parseInt(Utilities.formatDate(now, TIMEZONE, "mm"), 10);
  const isLate = (hours > 10 || (hours === 10 && minutes > 0)) ? "Yes" : "No";

  let photoUrl = (payload.photoUrl || payload.photoBase64)
    ? saveBase64Image(payload.photoUrl || payload.photoBase64, "Agent_" + (payload.agentName || "Selfie").replace(/\s+/g, "_"))
    : "";

  sheet.appendRow([
    dateStr,                                        // A: Date
    timeStr,                                        // B: Time
    payload.agentName || "",                        // C: Agent Name
    payload.leaderName || "",                       // D: Leader Name
    payload.typedLocation || payload.location || "",// E: Location (manual)
    photoUrl,                                       // F: Photo URL
    isLate,                                         // G: Is Late
    payload.gps || "",                              // H: GPS
    payload.resolvedLocation || "",                 // I: Location Detail
    payload.taskNotes || payload.notes || ""        // J: Notes
  ]);
  clearCache();
  return createJsonResponse({ status: "success", message: "Sign-on recorded.", isLate: isLate, photoUrl: photoUrl });
}

function handleLeaderCheckIn(ss, payload) {
  const sheet = getOrCreateSheet(ss, "LeaderCheckIns", [
    "Date", "Time", "Leader Name", "Region", "Odometer/Vehicle Info", "GPS", "Selfie/Photo URL", "Location"
  ]);
  const now = new Date();
  const dateStr = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, TIMEZONE, "HH:mm:ss");

  let vehiclePhotoUrl = payload.vehiclePhotoBase64
    ? saveBase64Image(payload.vehiclePhotoBase64, "Leader_Veh_" + (payload.leaderName || "Leader").replace(/\s+/g, "_"))
    : (payload.vehicleInfo || "");
  let selfiePhotoUrl = payload.selfiePhotoBase64
    ? saveBase64Image(payload.selfiePhotoBase64, "Leader_Selfie_" + (payload.leaderName || "Leader").replace(/\s+/g, "_"))
    : (payload.photoUrl || "");

  sheet.appendRow([
    dateStr,                                        // A: Date
    timeStr,                                        // B: Time
    payload.leaderName || "",                       // C: Leader Name
    payload.region || "",                           // D: Region
    vehiclePhotoUrl,                                // E: Odometer/Vehicle Info
    payload.gps || "",                              // F: GPS
    selfiePhotoUrl,                                 // G: Selfie/Photo URL
    payload.location || payload.region || ""        // H: Location
  ]);
  clearCache();
  return createJsonResponse({ status: "success", message: "Check-in recorded.", vehiclePhotoUrl, selfiePhotoUrl });
}

function handleSetWeeklyObjective(ss, payload) {
  const sheet = getOrCreateSheet(ss, "Objectives", [
    "Date", "View Type", "Target Date", "Leader Name",
    "Assigned Location", "Focus Areas", "Target Metrics", "Status", "Requested By",
    "Target Houses"
  ]);
  sheet.appendRow([
    Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd"),
    payload.viewType || "Weekly",
    payload.targetDate || Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd"),
    payload.leaderName || "",
    payload.assignedLocation || "",
    payload.focusAreas || "",
    payload.targetMetrics || "Active Sales",
    payload.status || "Published",
    payload.requestedBy || "Admin",
    payload.targetHouses || ""
  ]);
  clearCache();
  return createJsonResponse({ status: "success", message: "Objective published." });
}

function handleRequestObjective(ss, payload) {
  const sheet = getOrCreateSheet(ss, "Objectives", [
    "Date", "View Type", "Target Date", "Leader Name",
    "Assigned Location", "Focus Areas", "Target Metrics", "Status", "Requested By",
    "Target Houses"
  ]);
  sheet.appendRow([
    Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd"),
    payload.viewType || "Weekly",
    payload.targetDate || "",
    payload.leaderName || "",
    payload.assignedLocation || "",
    payload.focusAreas || "",
    payload.targetMetrics || "Active Sales",
    "Pending",
    payload.leaderName || "Team Leader",
    payload.targetHouses || ""
  ]);
  clearCache();
  return createJsonResponse({ status: "success", message: "Objective request submitted for Admin review." });
}

function handleApproveObjective(ss, payload) {
  const sheet = ss.getSheetByName("Objectives") || ss.getSheetByName("WeeklyObjectives");
  if (!sheet) return createJsonResponse({ status: "error", message: "Objectives sheet not found." });
  const rowIndex = parseInt(payload.rowIndex);
  if (!rowIndex || rowIndex < 2) {
    return createJsonResponse({ status: "error", message: "Invalid row index." });
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let statusCol = headers.indexOf("Status") + 1;
  if (statusCol <= 0) {
    statusCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, statusCol).setValue("Status").setFontWeight("bold").setBackground("#e2e8f0");
  }

  sheet.getRange(rowIndex, statusCol).setValue(payload.status || "Published");
  clearCache();
  return createJsonResponse({ status: "success", message: `Objective ${payload.status}.` });
}

function handleLogAgentStats(ss, payload) {
  const sheet = getOrCreateSheet(ss, "AgentStats", [
    "Date", "Agent Name", "Engagements", "Real Leads", "Payments",
    "Consecutive Misses (Strikes)", "Time", "Leader Name", "Issue Type", "Notes"
  ]);
  const now = new Date();
  const dateStr = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, TIMEZONE, "HH:mm:ss");

  sheet.appendRow([
    dateStr,
    payload.agentName || "",
    payload.engagements || "",
    payload.realLeads || "",
    payload.payments || "",
    payload.strikes || 0,
    timeStr,
    payload.leaderName || "",
    payload.issueType || (payload.strikes ? `${payload.strikes} Strike(s)` : "Performance Note"),
    payload.notes || ""
  ]);
  clearCache();
  return createJsonResponse({ status: "success", message: "Performance record saved." });
}

// ════════════════════════════════════════════════════════════════════════════
// 9. CACHE UTILITIES
// ════════════════════════════════════════════════════════════════════════════
function clearCache() {
  try { 
    CacheService.getScriptCache().remove("super_admin_portal_data"); 
    CacheService.getScriptCache().remove("super_admin_data"); 
  } catch (e) { }
}

// ════════════════════════════════════════════════════════════════════════════
// 10. DATA RETRIEVAL (GET PORTAL DATA)
// ════════════════════════════════════════════════════════════════════════════
function getSuperAdminData(ss) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("super_admin_portal_data");
  if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

  // ── 1. Agent Sign-Ons ──
  const agents = [];
  const agentSheet = ss.getSheetByName("AgentSignOns");
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const data = agentSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[2] && !row[0]) continue;
      agents.push({
        date: row[0],
        time: row[1],
        agentName: row[2],
        leaderName: row[3],
        location: row[4],
        photoUrl: row[5],
        isLate: row[6],
        gps: row[7],
        notes: row[9] || row[8] || ""
      });
    }
  }

  // ── 2. Leader Check-Ins ──
  const leaders = [];
  const leaderSheet = ss.getSheetByName("LeaderCheckIns");
  if (leaderSheet && leaderSheet.getLastRow() > 1) {
    const data = leaderSheet.getDataRange().getValues();
    const headers = data[0];
    const idx = (name) => { 
      const n = String(name).trim().toLowerCase();
      const i = headers.findIndex(h => String(h).trim().toLowerCase() === n);
      return i >= 0 ? i : null; 
    };
    const iDate = idx("Date") ?? 0;
    const iTime = idx("Time") ?? 1;
    const iName = idx("Leader Name") ?? 2;
    const iRegion = idx("Region") ?? 3;
    const iVeh = idx("Odometer/Vehicle Info") ?? (idx("Vehicle Info") ?? 4);
    const iGPS = idx("GPS") ?? 5;
    const iPhoto = idx("Selfie/Photo URL") ?? (idx("Photo URL") ?? 6);
    const iLocation = idx("Location") ?? 7;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[iName] && !row[0]) continue;
      leaders.push({
        date: row[iDate],
        time: row[iTime],
        leaderName: row[iName],
        region: row[iRegion],
        vehicleInfo: row[iVeh],
        photoUrl: row[iPhoto],
        gps: row[iGPS],
        location: row[iLocation] || row[iRegion] || ""
      });
    }
  }

  // ── 3. Objectives ──
  const objectives = [];
  const objSheet = ss.getSheetByName("Objectives") || ss.getSheetByName("WeeklyObjectives");
  if (objSheet && objSheet.getLastRow() > 1) {
    const data = objSheet.getDataRange().getValues();
    const headers = data[0];
    const idx = (name) => { const i = headers.indexOf(name); return i >= 0 ? i : null; };
    const iDate = idx("Date") ?? 0;
    const iViewType = idx("View Type") ?? 1;
    const iTargetDate = idx("Target Date") ?? 2;
    const iLeader = idx("Leader Name") ?? 3;
    const iLocation = idx("Assigned Location") ?? 4;
    const iFocus = idx("Focus Areas") ?? 5;
    const iMetrics = idx("Target Metrics") ?? 6;
    const iStatus = idx("Status") ?? 7;
    const iReqBy = idx("Requested By") ?? 8;
    const iHouses = idx("Target Houses") ?? 9;

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      objectives.push({
        rowIndex: i + 1,
        date: r[iDate],
        viewType: r[iViewType],
        targetDate: r[iTargetDate],
        leaderName: r[iLeader],
        assignedLocation: r[iLocation],
        focusAreas: r[iFocus],
        targetMetrics: r[iMetrics],
        status: (iStatus !== null && r[iStatus]) ? r[iStatus] : "Published",
        requestedBy: (iReqBy !== null && r[iReqBy]) ? r[iReqBy] : "Admin",
        targetHouses: (iHouses !== null && r[iHouses]) ? r[iHouses] : ""
      });
    }
  }

  // ── 4. Agent Performance & Strikes (AgentStats) ──
  const agentStats = [];
  const perfSheet = ss.getSheetByName("AgentStats") || ss.getSheetByName("AgentPerformance");
  if (perfSheet && perfSheet.getLastRow() > 1) {
    const data = perfSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[1] && !r[0]) continue;
      agentStats.push({
        date: r[0],
        agentName: r[1],
        engagements: r[2] || "",
        realLeads: r[3] || "",
        payments: r[4] || "",
        strikes: r[5] || 0,
        time: r[6] || "",
        leaderName: r[7] || "",
        issueType: r[8] || (r[5] ? `${r[5]} Strike(s)` : "Performance Note"),
        notes: r[9] || r[3] || ""
      });
    }
  }

  // ── 5. Team Leaders List from Users sheet ──
  const teamLeadersList = [];
  const usersSheet = ss.getSheetByName("Users");
  if (usersSheet && usersSheet.getLastRow() > 1) {
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0];
    const idx = (name) => { const i = headers.indexOf(name); return i >= 0 ? i : null; };
    const iEmail = idx("Email") ?? 0;
    const iFirstName = idx("First Name") ?? 1;
    const iLastName = idx("Last Name") ?? 2;
    const iRole = idx("Account Type") ?? 5;

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (String(r[iRole]).trim().toLowerCase() === "team leader") {
        teamLeadersList.push({
          firstName: r[iFirstName],
          lastName: r[iLastName],
          email: r[iEmail],
          fullName: `${r[iFirstName]} ${r[iLastName]}`.trim()
        });
      }
    }
  }

  const outputPayload = {
    status: "success",
    agents: agents.reverse(),
    leaders: leaders.reverse(),
    objectives: objectives.reverse(),
    agentStats: agentStats.reverse(),
    teamLeadersList: teamLeadersList
  };

  const jsonStr = JSON.stringify(outputPayload);
  if (jsonStr.length < 95000) {
    try { cache.put("super_admin_portal_data", jsonStr, 180); } catch (e) { }
  }
  return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
}

function getObjectivesDataHandler(ss) {
  const result = getSuperAdminData(ss);
  return result;
}

