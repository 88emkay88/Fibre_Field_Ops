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
// 1. GET REQUEST ROUTER & CACHE
// ════════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : null;
    const ss = getSpreadsheetInstance();

    if (action === "getSuperAdminData" || action === "getPortalData" || action === "getLeaderData" || action === "getAgentData") {
      const cache = CacheService.getScriptCache();
      const cached = cache.get("super_admin_portal_data");
      
      if (cached) {
        return ContentService.createTextOutput(cached)
          .setMimeType(ContentService.MimeType.JSON);
      }

      const result = getSuperAdminData(ss);
      const jsonStr = JSON.stringify(result);
      
      // Cache payload if within size limit (100KB limit for Apps Script cache)
      if (jsonStr.length < 95000) {
        try {
          cache.put("super_admin_portal_data", jsonStr, 120); // 2 minutes
        } catch (cErr) {
          Logger.log("Cache put skipped: " + cErr.toString());
        }
      }
      
      return ContentService.createTextOutput(jsonStr)
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return createJsonResponse({ status: "success", message: "FibreGems Operations API is active." });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 2. POST REQUEST ROUTER
// ════════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "No post data received." });
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload || {};
    const ss = getSpreadsheetInstance();

    Logger.log(`Action: ${action}`);

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

      case "deleteAgentStat":
        return createJsonResponse(deleteAgentStat(ss, payload));

      // ── Geocoding ──
      case "getReadableLocation":
        return createJsonResponse(getReadableLocationBackend(payload.lat, payload.lon));

      default:
        return createJsonResponse({ status: "error", message: "Invalid action: " + action });
    }
  } catch (err) {
    Logger.log("doPost Error: " + err.toString());
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 3. ONE-CLICK SCHEMA MIGRATION SCRIPT
// ════════════════════════════════════════════════════════════════════════════
/**
 * Run this function once from the Apps Script toolbar to automatically
 * add missing columns to your existing sheets without losing any data.
 */
function migrateAndExtendSheets() {
  const ss = getSpreadsheetInstance();
  if (!ss) {
    Logger.log("❌ Error: Could not connect to spreadsheet.");
    return;
  }

  // 1. Objectives Sheet
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
    "Date", "Time", "Leader Name", "Region", "Vehicle Info", "Photo URL", "GPS", "Location"
  ]);

  // 5. Users Sheet
  extendSheetHeaders(ss, "Users", [
    "Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Reset Token"
  ]);

  clearCache();
  Logger.log("✅ All sheets checked and successfully extended!");
}

function extendSheetHeaders(ss, sheetName, expectedHeaders) {
  if (!ss) ss = getSpreadsheetInstance();
  if (!ss) return;

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(expectedHeaders);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight("bold");
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
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]).setFontWeight("bold");
    Logger.log(`Extended ${sheetName} with columns: ${missingHeaders.join(", ")}`);
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 4. AUTHENTICATION HANDLERS
// ════════════════════════════════════════════════════════════════════════════
function loginUser(ss, email, password) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && String(row[0]).trim().toLowerCase() === emailLower) {
        const storedHash = row[3];
        const salt = row[4];
        const testHash = computeSHA256(password, salt);

        if (storedHash === testHash) {
          return {
            status: "success",
            user: {
              email: row[0],
              firstName: row[1],
              lastName: row[2],
              accountType: row[5] || "Agent",
            }
          };
        } else {
          return { status: "error", message: "Invalid password." };
        }
      }
    }
    return { status: "error", message: "No user found with this email." };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function registerUser(ss, email, firstName, lastName, password) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === emailLower) {
        return { status: "error", message: "A user with this email already exists." };
      }
    }

    const salt = generateSalt(16);
    const hash = computeSHA256(password, salt);

    sheet.appendRow([
      emailLower,
      firstName,
      lastName,
      hash,
      salt,
      "Agent",
      "Active",
      ""
    ]);

    clearCache();
    return {
      status: "success",
      user: { email: emailLower, firstName, lastName, accountType: "Agent" }
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function addTeamLeader(ss, firstName, lastName, email, password, accountType) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === emailLower) {
        return { status: "error", message: "A user with this email already exists." };
      }
    }

    if (!firstName || !lastName || !email || !password) {
      return { status: "error", message: "All fields are required." };
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
    return { status: "success", message: `Team Leader ${firstName} ${lastName} added successfully.` };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function deleteTeamLeader(ss, email) {
  try {
    const sheet = getOrCreateUsersSheet(ss);
    const data = sheet.getDataRange().getValues();
    const emailLower = String(email || "").trim().toLowerCase();

    if (!emailLower) return { status: "error", message: "Email is required." };

    let deleted = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === emailLower) {
        sheet.deleteRow(i + 1);
        deleted = true;
        break;
      }
    }

    if (!deleted) return { status: "error", message: "Team Leader not found." };

    clearCache();
    return { status: "success", message: `Team Leader ${email} removed successfully.` };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 5. FIELD SIGN-ON & CHECK-IN HANDLERS
// ════════════════════════════════════════════════════════════════════════════
function handleAgentSignOn(ss, payload) {
  let sheet = ss.getSheetByName("AgentSignOns");
  if (!sheet) {
    migrateAndExtendSheets();
    sheet = ss.getSheetByName("AgentSignOns");
  }

  const now = new Date();
  const timeStr = formatTime(now);
  const hour = now.getHours();
  const minute = now.getMinutes();
  const isLate = (hour > 10 || (hour === 10 && minute > 0)) ? "Yes" : "No";

  const locationDisplay = payload.resolvedLocation || payload.typedLocation || payload.location || "";
  const locationDetail = payload.typedLocation || payload.location || "";

  sheet.appendRow([
    now,
    timeStr,
    payload.agentName || "",
    payload.leaderName || "",
    locationDisplay,
    payload.photoUrl || "",
    isLate,
    payload.gps || "",
    locationDetail,
    payload.taskNotes || payload.notes || ""
  ]);

  clearCache();
  return createJsonResponse({ status: "success", isLate: isLate, message: "Agent sign-on recorded." });
}

function handleLeaderCheckIn(ss, payload) {
  let sheet = ss.getSheetByName("LeaderCheckIns");
  if (!sheet) {
    migrateAndExtendSheets();
    sheet = ss.getSheetByName("LeaderCheckIns");
  }

  const now = new Date();
  const timeStr = formatTime(now);

  sheet.appendRow([
    now,
    timeStr,
    payload.leaderName || "",
    payload.region || "",
    payload.vehiclePhotoBase64 || payload.vehicleInfo || "",
    payload.selfiePhotoBase64 || payload.photoUrl || "",
    payload.gps || "",
    payload.location || payload.region || ""
  ]);

  clearCache();
  return createJsonResponse({ status: "success", message: "Leader check-in recorded." });
}

// ════════════════════════════════════════════════════════════════════════════
// 6. OBJECTIVES HANDLERS
// ════════════════════════════════════════════════════════════════════════════
function handleSetWeeklyObjective(ss, payload) {
  const sheetName = ss.getSheetByName("Objectives") ? "Objectives" : "WeeklyObjectives";
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    migrateAndExtendSheets();
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.appendRow([
    new Date(),
    payload.viewType || "Week",
    payload.targetDate || "",
    payload.leaderName || "",
    payload.assignedLocation || "",
    payload.focusAreas || "",
    payload.targetMetrics || "",
    payload.status || "Published",
    payload.requestedBy || "Admin",
    payload.targetHouses || ""
  ]);

  clearCache();
  return createJsonResponse({ status: "success", message: "Objective published." });
}

function handleRequestObjective(ss, payload) {
  const sheetName = ss.getSheetByName("Objectives") ? "Objectives" : "WeeklyObjectives";
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    migrateAndExtendSheets();
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.appendRow([
    new Date(),
    payload.viewType || "Week",
    payload.targetDate || "",
    payload.leaderName || "",
    payload.assignedLocation || "",
    payload.focusAreas || "",
    payload.targetMetrics || "",
    "Pending",
    payload.leaderName || "Team Leader",
    payload.targetHouses || ""
  ]);

  clearCache();
  return createJsonResponse({ status: "success", message: "Objective request submitted." });
}

function handleApproveObjective(ss, payload) {
  const sheetName = ss.getSheetByName("Objectives") ? "Objectives" : "WeeklyObjectives";
  const sheet = ss.getSheetByName(sheetName);
  const rowIndex = parseInt(payload.rowIndex);
  const status = payload.status || "Published";

  if (!sheet || !rowIndex || rowIndex < 2) {
    return createJsonResponse({ status: "error", message: "Invalid objective index." });
  }

  // Status is column 8
  sheet.getRange(rowIndex, 8).setValue(status);
  clearCache();
  return createJsonResponse({ status: "success", message: `Objective status set to ${status}.` });
}

// ════════════════════════════════════════════════════════════════════════════
// 7. AGENT PERFORMANCE & STRIKE HANDLER
// ════════════════════════════════════════════════════════════════════════════
function handleLogAgentStats(ss, payload) {
  let sheet = ss.getSheetByName("AgentStats");
  if (!sheet) {
    migrateAndExtendSheets();
    sheet = ss.getSheetByName("AgentStats");
  }

  const now = new Date();
  sheet.appendRow([
    now,
    payload.agentName || "",
    payload.engagements || "",
    payload.realLeads || "",
    payload.payments || "",
    payload.strikes || 0,
    formatTime(now),
    payload.leaderName || "",
    payload.issueType || "Performance Note",
    payload.notes || ""
  ]);

  clearCache();
  return createJsonResponse({ status: "success", message: "Performance record saved." });
}

function deleteAgentStat(ss, payload) {
  try {
    const sheet = ss.getSheetByName("AgentStats") || ss.getSheetByName("AgentPerformance");
    if (!sheet) return { status: "error", message: "AgentStats sheet not found." };

    const rowIndex = parseInt(payload.rowIndex);
    const agentName = String(payload.agentName || "").trim().toLowerCase();
    const clearAll = payload.clearAll === true;
    const data = sheet.getDataRange().getValues();

    let deletedCount = 0;
    if (clearAll && agentName) {
      // Delete all matching rows for this agent
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]).trim().toLowerCase() === agentName) {
          sheet.deleteRow(i + 1);
          deletedCount++;
        }
      }
    } else if (rowIndex && rowIndex >= 2 && rowIndex <= data.length) {
      sheet.deleteRow(rowIndex);
      deletedCount++;
    } else if (agentName) {
      // Delete most recent row for this agent
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]).trim().toLowerCase() === agentName) {
          sheet.deleteRow(i + 1);
          deletedCount++;
          break;
        }
      }
    }

    if (deletedCount === 0) return { status: "error", message: "No matching record found to delete." };

    clearCache();
    return { status: "success", message: `Deleted ${deletedCount} record(s) successfully.` };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════════════════════
// 8. DATA RETRIEVAL (GET PORTAL DATA)
// ════════════════════════════════════════════════════════════════════════════
function getSuperAdminData(ss) {
  // 1. Objectives
  const objSheet = ss.getSheetByName("Objectives") || ss.getSheetByName("WeeklyObjectives");
  const objectives = [];
  if (objSheet && objSheet.getLastRow() > 1) {
    const rows = objSheet.getRange(2, 1, objSheet.getLastRow() - 1, objSheet.getLastColumn()).getValues();
    rows.forEach((r, idx) => {
      objectives.push({
        rowIndex: idx + 2,
        date: r[0],
        viewType: r[1],
        targetDate: r[2],
        leaderName: r[3],
        assignedLocation: r[4],
        focusAreas: r[5],
        targetMetrics: r[6],
        status: r[7] || "Published",
        requestedBy: r[8] || "",
        targetHouses: r[9] || ""
      });
    });
  }

  // 2. Agent Stats / Strikes
  const statSheet = ss.getSheetByName("AgentStats");
  const agentStats = [];
  if (statSheet && statSheet.getLastRow() > 1) {
    const rows = statSheet.getRange(2, 1, statSheet.getLastRow() - 1, statSheet.getLastColumn()).getValues();
    rows.forEach((r, idx) => {
      agentStats.push({
        rowIndex: idx + 2,
        date: r[0],
        agentName: r[1],
        engagements: r[2],
        realLeads: r[3],
        payments: r[4],
        strikes: r[5] || 0,
        time: r[6] || "",
        leaderName: r[7] || "",
        issueType: r[8] || (r[5] ? `${r[5]} Strike(s)` : "Performance Note"),
        notes: r[9] || ""
      });
    });
  }

  // 3. Team Leaders List from Users sheet
  const usersSheet = ss.getSheetByName("Users");
  const teamLeadersList = [];
  if (usersSheet && usersSheet.getLastRow() > 1) {
    const uRows = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, usersSheet.getLastColumn()).getValues();
    uRows.forEach(r => {
      if (String(r[5]).toLowerCase() === "team leader") {
        teamLeadersList.push({
          email: r[0],
          firstName: r[1],
          lastName: r[2],
          fullName: `${r[1]} ${r[2]}`.trim()
        });
      }
    });
  }

  return {
    status: "success",
    agents: getAgentDataRows(ss),
    leaders: getLeaderDataRows(ss),
    objectives: objectives.reverse(),
    agentStats: agentStats.reverse(),
    teamLeadersList: teamLeadersList
  };
}

function getAgentDataRows(ss) {
  const sheet = ss.getSheetByName("AgentSignOns");
  const agents = [];
  if (sheet && sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    rows.forEach(r => {
      agents.push({
        date: r[0],
        time: r[1],
        agentName: r[2],
        leaderName: r[3],
        location: r[4],
        photoUrl: r[5],
        isLate: r[6],
        gps: r[7],
        notes: r[9] || r[8] || ""
      });
    });
  }
  return agents.reverse();
}

function getLeaderDataRows(ss) {
  const sheet = ss.getSheetByName("LeaderCheckIns");
  const leaders = [];
  if (sheet && sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    rows.forEach(r => {
      leaders.push({
        date: r[0],
        time: r[1],
        leaderName: r[2],
        region: r[3],
        vehicleInfo: r[4],
        photoUrl: r[5],
        gps: r[6],
        location: r[7] || r[3] || ""
      });
    });
  }
  return leaders.reverse();
}

// ════════════════════════════════════════════════════════════════════════════
// 9. UTILITIES & HELPERS
// ════════════════════════════════════════════════════════════════════════════
function getSpreadsheetInstance() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss;
}

function getOrCreateUsersSheet(ss) {
  let sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Reset Token"]);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
  }
  return sheet;
}

function getReadableLocationBackend(lat, lon) {
  try {
    const apiKey = "pk.1be07ee2080691339d8fc4f1712dbc95";
    const url = `https://us1.locationiq.com/v1/reverse?key=${apiKey}&lat=${lat}&lon=${lon}&format=json`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data && data.address) {
      const addr = data.address;
      const localArea = addr.suburb || addr.neighbourhood || addr.residential || "";
      const city = addr.city || addr.town || addr.village || addr.county || "";
      let parts = [];
      if (localArea) parts.push(localArea);
      if (city && city !== localArea) parts.push(city);

      if (parts.length > 0) return { status: "success", address: parts.join(", ") };
      return { status: "success", address: data.display_name.split(",").slice(0, 3).join(",") };
    }
    return { status: "error", message: "Location not found" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clearCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove("super_admin_portal_data");
  } catch (e) {}
}

function formatTime(d) {
  if (!d) return "";
  const date = new Date(d);
  const hours = date.getHours().toString().padStart(2, "0");
  const minutes = date.getMinutes().toString().padStart(2, "0");
  return `${hours}:${minutes}`;
}

function computeSHA256(password, salt) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + salt,
    Utilities.Charset.UTF_8
  );
  return rawHash
    .map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, "0"))
    .join("");
}

function generateSalt(length) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}
