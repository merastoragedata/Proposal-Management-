// ═══════════════════════════════════════════════════════════════
// MSETCL Estimate Generator – Google Apps Script Backend v5
// ═══════════════════════════════════════════════════════════════
// DEPLOY: Extensions → Apps Script → Deploy → New Deployment
//         Type: Web App | Execute as: Me | Access: Anyone
//
// MIGRATION NOTE (v4 → v5): This version ADDS columns to existing
// sheets via getOrCreateSheet's additive-column logic. It does NOT
// delete or rename any existing column, so all current SOR items,
// users, and saved estimates are preserved untouched. New columns
// added to Estimates: estimateType, supplySpares, amcAmount,
// amcRemark, viewerVisible, progress (JSON: stages/financial/
// orderClosed/images), supplyTotal, projectTotal, omTotal,
// nonSorSupplyTotal, nonSorOmTotal. The legacy "sentForApproval"
// and "approvalRemarks" columns remain in the sheet (untouched, for
// safety) but are no longer read or written by this script — the
// 25-stage checklist replaces them, stored inside the "progress"
// JSON column. The TCodes/Processes sheets gain a "status" column
// (approved|pending) — existing rows have no value, treated as
// "approved" by getTCodes/getProcesses for backward compatibility.
// ═══════════════════════════════════════════════════════════════

// BACKEND_VERSION — bump this whenever the action set changes. Surfaced by
// ping() so the Admin Panel's "Test Connection" can prove whether the
// deployment you're actually hitting has the latest actions (e.g. image
// upload) or whether you're calling a stale/different deployment URL.
var BACKEND_VERSION = "v8-images-2026-06-21";

var SS_NAME = "MSETCL Estimate Generator DB";

// ── PIN TO A SPECIFIC SPREADSHEET (recommended) ──────────────────
// Paste your existing/production spreadsheet's ID below to guarantee this
// script always uses that exact file, regardless of name or which Google
// account owns it relative to whoever deployed the script. Find the ID in
// the spreadsheet's URL: https://docs.google.com/spreadsheets/d/<THIS_PART>/edit
// Leave as "" to fall back to find-or-create-by-name (the old, fragile
// behavior — only safe if you're certain only one account/file is in play).
var FIXED_SS_ID = "";
var SHEET_SOR       = "SOR_Database";
var SHEET_ESTIMATES = "Estimates";
var SHEET_ITEMS     = "Estimate_Items";
var SHEET_USERS     = "Users";
var SHEET_NONSOR    = "NonSOR_Items";
var SHEET_TCODES    = "TCodes";
var SHEET_PROCESSES = "Processes";

var SOR_HDR    = ["item_no","description","uom","rate","section","version"];

// EST_HDR: additive — old columns kept in original order/position,
// new columns appended at the end so existing rows stay valid.
var EST_HDR    = ["id","name","number","date","division","ref","work",
                  "sorVersions","addCharge","settings","finalTotal",
                  "savedAt","status","itemCount","ownerId","ownerName",
                  "audited","sentForApproval","approvalRemarks",
                  "remark1","remark2","remark3","remark4","remark5",
                  "viewerVisible","estimateType","amcAmount","progressImages",
                  // ── new in v5 ──
                  "supplySpares","amcRemark","progress",
                  "supplyTotal","projectTotal","omTotal",
                  "nonSorSupplyTotal","nonSorOmTotal",
                  // ── new in v6 ──
                  "trLineMode","landCompChargePct","centagesOverridden",
                  // ── new in v7 ──
                  "pmoPrEntries",
                  // ── new in v8 ──
                  "prSettings"];

var ITEM_HDR   = ["estimate_id","item_no","version","description",
                  "uom","rate","qty","section","remarks","isCustom"];
var USER_HDR   = ["id","name","passHash","role","status","createdAt"];
var NONSOR_HDR = ["item_no","description","uom","rate","section","remarks","addedBy","addedAt"];
var TCODE_HDR  = ["tcode","purpose","description","remarks","addedBy","addedAt","status","image"];
var PROC_HDR   = ["name","description","remarks","addedBy","addedAt","status","image"];

// ── CORS helper ─────────────────────────────────────────────────
function cors(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ────────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "ping";
  var result;
  try {
    if      (action === "ping")           result = ping();
    else if (action === "init")           result = initDB();
    else if (action === "getSOR")         result = getSOR(e.parameter.version || "");
    else if (action === "getSORVersions") result = getSORVersions();
    else if (action === "getEstimates")   result = getEstimates(e.parameter.userId||"", e.parameter.role||"");
    else if (action === "getEstimate")    result = getEstimate(e.parameter.id || "", e.parameter.userId||"", e.parameter.role||"");
    else if (action === "getUsers")       result = getUsers();
    else if (action === "getNonSOR")      result = getNonSOR();
    else if (action === "getTCodes")      result = getTCodes();
    else if (action === "getProcesses")   result = getProcesses();
    else if (action === "getChargeRates") result = getChargeRates();
    else if (action === "bootstrap")      result = bootstrap(e.parameter.userId||"", e.parameter.role||"");
    else if (action === "getAnthropicApiKeyStatus") result = getAnthropicApiKeyStatus();
    else if (action === "getAiAllowedUsers")        result = getAiAllowedUsers();
    else result = {ok:false, error:"Unknown action: "+action};
  } catch(err) {
    result = {ok:false, error:err.toString(), stack: err.stack||""};
  }
  return cors(result);
}

// ── POST ───────────────────────────────────────────────────────
function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || "";
    if      (action === "saveEstimate")    result = saveEstimate(body.estimate, body.items);
    else if (action === "deleteEstimate")  result = deleteEstimate(body.id, body.userId||"", body.role||"");
    else if (action === "importSOR")       result = importSOR(body.items, body.version);
    else if (action === "registerUser")    result = registerUser(body.user);
    else if (action === "loginUser")       result = loginUser(body.id, body.passHash);
    else if (action === "saveUser")        result = saveUser(body.user, body.actingRole||"");
    else if (action === "deleteUser")      result = deleteUser(body.id, body.actingRole||"");
    else if (action === "approveUser")     result = approveUser(body.id, body.approve, body.actingRole||"");
    else if (action === "changePassword")  result = changePassword(body.id, body.oldHash, body.newHash);
    else if (action === "addNonSOR")       result = addNonSOR(body.item);
    else if (action === "updateNonSOR")    result = updateNonSOR(body.item, body.actingRole||"");
    else if (action === "deleteNonSOR")    result = deleteNonSOR(body.item_no, body.actingRole||"");
    else if (action === "addTCode")        result = addTCode(body.item);
    else if (action === "updateTCode")     result = updateTCode(body.item, body.userId||"", body.role||"");
    else if (action === "deleteTCode")     result = deleteTCode(body.tcode, body.userId||"", body.role||"");
    else if (action === "addProcess")      result = addProcess(body.item);
    else if (action === "updateProcess")   result = updateProcess(body.item, body.userId||"", body.role||"");
    else if (action === "deleteProcess")   result = deleteProcess(body.name, body.userId||"", body.role||"");
    else if (action === "resetChargeRates")result = resetChargeRates();
    else if (action === "saveChargeRates") result = saveChargeRates(body.rates, body.actingRole||"");
    else if (action === "uploadEstimateImage") result = uploadEstimateImage(body.estimateId, body.estimateName, body.fileName, body.mimeType, body.base64Data);
    else if (action === "deleteEstimateImage") result = deleteEstimateImage(body.fileId);
    else if (action === "uploadTcpImage")  result = uploadTcpImage(body.kind, body.key, body.fileName, body.mimeType, body.base64Data);
    else if (action === "saveAnthropicApiKey") result = saveAnthropicApiKey(body.apiKey, body.actingRole||"");
    else if (action === "setAiAllowedUsers")   result = setAiAllowedUsers(body.userIds, body.actingRole||"");
    else if (action === "aiCreateEstimateDraft") result = aiCreateEstimateDraft(body.userId||"", body.role||"", body.requestText, body.sorExcerpt||"", body.nonSorExcerpt||"");
    else result = {ok:false, error:"Unknown action: "+action};
  } catch(err) {
    result = {ok:false, error:err.toString()};
  }
  return cors(result);
}

// ── SPREADSHEET HELPERS ────────────────────────────────────────
function getOrCreateSS() {
  if (FIXED_SS_ID) {
    return SpreadsheetApp.openById(FIXED_SS_ID);
  }
  var files = DriveApp.getFilesByName(SS_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(SS_NAME);
}

// Additive: never removes/renames columns. Appends any missing
// trailing headers to existing sheets so old rows stay valid.
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground("#0F2044")
         .setFontColor("#FFFFFF")
         .setFontWeight("bold");
  } else if (sheet.getLastColumn() < headers.length) {
    var curHeaders = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    for (var i = curHeaders.length; i < headers.length; i++) {
      sheet.getRange(1, i+1).setValue(headers[i])
           .setBackground("#0F2044").setFontColor("#FFFFFF").setFontWeight("bold");
    }
  }
  return sheet;
}

function tryParse(val, def) {
  if (val === null || val === undefined || val === "") return def;
  try { return JSON.parse(val); } catch(e) { return def; }
}

// ── DRIVE IMAGE STORAGE ──────────────────────────────────────────
// Progress images and T-Code/Process images are stored as real Drive files
// (not base64-in-sheet) under a root folder, with one subfolder per
// estimate created lazily on first image upload for that estimate.
var DRIVE_ROOT_FOLDER_NAME = "MSETCL Estimate Generator - Images";

function getOrCreateDriveRootFolder_() {
  var folders = DriveApp.getFoldersByName(DRIVE_ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_ROOT_FOLDER_NAME);
}
function getOrCreateEstimateFolder_(estimateId, estimateName) {
  var root = getOrCreateDriveRootFolder_();
  var folderName = estimateName ? (estimateName + " [" + estimateId + "]") : estimateId;
  var existing = root.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return root.createFolder(folderName);
}
function getOrCreateTcpFolder_() {
  var root = getOrCreateDriveRootFolder_();
  var name = "TCode_Process_Images";
  var existing = root.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return root.createFolder(name);
}

function uploadEstimateImage(estimateId, estimateName, fileName, mimeType, base64Data) {
  if (!estimateId) return {ok:false, error:"Missing estimate ID — save the estimate before adding images"};
  if (!base64Data) return {ok:false, error:"No image data received"};
  try {
    var folder = getOrCreateEstimateFolder_(estimateId, estimateName);
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType||"image/jpeg", fileName||("image_"+new Date().getTime()+".jpg"));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {ok:true, fileId: file.getId(), url: "https://drive.google.com/uc?id="+file.getId(), viewUrl: file.getUrl()};
  } catch(err) {
    return {ok:false, error:"Drive upload failed: "+err.toString()};
  }
}
function deleteEstimateImage(fileId) {
  if (!fileId) return {ok:false, error:"Missing file ID"};
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return {ok:true};
  } catch(err) {
    return {ok:false, error:"Delete failed: "+err.toString()};
  }
}
function uploadTcpImage(kind, key, fileName, mimeType, base64Data) {
  if (!key) return {ok:false, error:"Missing T Code / Process key"};
  if (!base64Data) return {ok:false, error:"No image data received"};
  try {
    var folder = getOrCreateTcpFolder_();
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType||"image/jpeg", fileName||((kind||"item")+"_"+key+".jpg"));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = "https://drive.google.com/uc?id="+file.getId();
    // Persist the URL directly onto the TCode/Process row
    var ss = getOrCreateSS();
    if (kind === "tcode") {
      var tSheet = getOrCreateSheet(ss, SHEET_TCODES, TCODE_HDR);
      var tRow = findTCodeRow_(tSheet, key);
      if (tRow > -1) tSheet.getRange(tRow, 8).setValue(url);
    } else {
      var pSheet = getOrCreateSheet(ss, SHEET_PROCESSES, PROC_HDR);
      var pRow = findProcessRow_(pSheet, key);
      if (pRow > -1) pSheet.getRange(pRow, 7).setValue(url);
    }
    return {ok:true, url: url, fileId: file.getId(), viewUrl: file.getUrl()};
  } catch(err) {
    return {ok:false, error:"Drive upload failed: "+err.toString()};
  }
}

// ── PING ───────────────────────────────────────────────────────
function ping() {
  try {
    var files = DriveApp.getFilesByName(SS_NAME);
    var exists = files.hasNext();
    var sorRows = 0, estRows = 0, ssUrl = "", ssId = "", ssName2 = SS_NAME;
    if (exists) {
      var ss = SpreadsheetApp.open(files.next());
      ssUrl  = ss.getUrl();
      ssId   = ss.getId();
      ssName2= ss.getName();
      var sorSheet = ss.getSheetByName(SHEET_SOR);
      var estSheet = ss.getSheetByName(SHEET_ESTIMATES);
      if (sorSheet) sorRows = Math.max(0, sorSheet.getLastRow() - 1);
      if (estSheet) estRows = Math.max(0, estSheet.getLastRow() - 1);
    }
    return {
      ok: true, ssExists: exists, ssId: ssId, ssUrl: ssUrl, ssName: ssName2,
      sorRows: sorRows, estRows: estRows, initialized: exists,
      backendVersion: BACKEND_VERSION, hasImageActions: true
    };
  } catch(err) {
    return {ok:false, error:err.toString()};
  }
}

// ── INIT ───────────────────────────────────────────────────────
function initDB() {
  var ss = getOrCreateSS();
  getOrCreateSheet(ss, SHEET_SOR,       SOR_HDR);
  getOrCreateSheet(ss, SHEET_ESTIMATES, EST_HDR);
  getOrCreateSheet(ss, SHEET_ITEMS,     ITEM_HDR);
  getOrCreateSheet(ss, SHEET_NONSOR,    NONSOR_HDR);
  getOrCreateSheet(ss, SHEET_TCODES,    TCODE_HDR);
  getOrCreateSheet(ss, SHEET_PROCESSES, PROC_HDR);
  var userSheet = getOrCreateSheet(ss, SHEET_USERS, USER_HDR);
  ensureSeedAdmin_(userSheet);
  var cfgSheet = ss.getSheetByName("Config");
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet("Config");
    cfgSheet.getRange(1,1,1,2).setValues([["key","value"]]);
    cfgSheet.getRange(1,1,1,2).setBackground("#0F2044").setFontColor("#FFFFFF").setFontWeight("bold");
    cfgSheet.appendRow(["chargeRates", JSON.stringify(defaultChargeRates())]);
  }
  var def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1) {
    try { ss.deleteSheet(def); } catch(e) {}
  }
  return {ok:true, ssUrl:ss.getUrl(), ssId:ss.getId(), message:"Initialized"};
}

function hashPassGS(p) {
  var h = 0;
  for (var i=0;i<p.length;i++){ h = ((h<<5)-h)+p.charCodeAt(i); h |= 0; }
  return 'h'+Math.abs(h).toString(36);
}

// ── CHARGE RATES (system defaults, admin-editable) ────────────
function defaultChargeRates() {
  return {additionalCharge:15, labourCess:1, freightInsurance:5, gst:18, pv:5, contingencies:3, centages:10};
}
function getConfigSheet_(ss) {
  var sheet = ss.getSheetByName("Config");
  if (!sheet) {
    sheet = ss.insertSheet("Config");
    sheet.getRange(1,1,1,2).setValues([["key","value"]]);
    sheet.getRange(1,1,1,2).setBackground("#0F2044").setFontColor("#FFFFFF").setFontWeight("bold");
  }
  return sheet;
}
// ── Generic config key/value read/write, reused for the Anthropic API key,
// the AI-estimate-creator allow-list, and any other single-value settings. ──
function getConfigValue_(key, def) {
  var ss = getOrCreateSS();
  var sheet = getConfigSheet_(ss);
  if (sheet.getLastRow() < 2) return def;
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
  for (var i=0;i<data.length;i++) if (data[i][0]===key) return data[i][1];
  return def;
}
function setConfigValue_(key, value) {
  var ss = getOrCreateSS();
  var sheet = getConfigSheet_(ss);
  var data = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues() : [];
  for (var i=0;i<data.length;i++) {
    if (data[i][0]===key) { sheet.getRange(i+2,2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

// ── AI Estimate Creator — admin sets a shared Anthropic API key once;
// approved users (admin allow-list) can use natural-language estimate
// creation. The key is NEVER returned to the frontend — only used
// server-side when calling the Anthropic API. ──
function saveAnthropicApiKey(apiKey, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can set the API key"};
  setConfigValue_("anthropicApiKey", apiKey||"");
  return {ok:true};
}
function getAnthropicApiKeyStatus() {
  var key = getConfigValue_("anthropicApiKey", "");
  return {ok:true, configured: !!key};
}
function getAiAllowedUsers() {
  var list = tryParse(getConfigValue_("aiAllowedUsers", "[]"), []);
  return {ok:true, users: list};
}
function setAiAllowedUsers(userIds, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can manage this list"};
  setConfigValue_("aiAllowedUsers", JSON.stringify(userIds||[]));
  return {ok:true};
}
function isUserAiAllowed_(userId, role) {
  if (role === "admin") return true;
  var list = tryParse(getConfigValue_("aiAllowedUsers", "[]"), []);
  return list.indexOf(userId) > -1;
}

// ── Calls Anthropic's Messages API with the user's natural-language
// estimate request plus a compact SOR/Non-SOR catalog excerpt, asking for
// a structured JSON item list back. Returns {ok, items:[{item_no,
// description, uom, rate, section, qty, source, remark}], explanation}. ──
function aiCreateEstimateDraft(userId, role, requestText, sorExcerpt, nonSorExcerpt) {
  if (!isUserAiAllowed_(userId, role)) return {ok:false, error:"You are not approved for AI Estimate Creator. Ask an admin to enable access."};
  var apiKey = getConfigValue_("anthropicApiKey", "");
  if (!apiKey) return {ok:false, error:"AI Estimate Creator is not configured yet. An admin must set the Anthropic API key in Admin Panel."};
  if (!requestText || !requestText.trim()) return {ok:false, error:"Describe what you need in the estimate first."};

  var systemPrompt = "You are an assistant that converts an MSETCL EHV substation engineer's natural-language estimate request into a structured list of line items, using ONLY the provided SOR (Schedule of Rates) and Non-SOR catalog excerpts. " +
    "For each item needed, pick the closest matching catalog entry and a reasonable quantity based on the request. " +
    "If something is needed but not in either catalog, invent a reasonable industry-standard description, UOM, and rate, and mark source as 'estimated' with a remark explaining it's not from the catalog. " +
    "Respond with ONLY a JSON object (no markdown, no prose) of the exact shape: " +
    '{"items":[{"item_no":"...","description":"...","uom":"...","rate":0,"section":"supply|project|om","qty":0,"source":"sor|nonsor|estimated","remark":""}],"explanation":"one paragraph summary of what was included and why"}';

  var userPrompt = "Request: " + requestText + "\n\nSOR catalog excerpt (item_no | description | uom | rate | section):\n" + sorExcerpt +
    "\n\nNon-SOR catalog excerpt (item_no | description | uom | rate | section):\n" + nonSorExcerpt;

  var payload = {
    model: "claude-sonnet-4-6",
    max_tokens: 4000,
    system: systemPrompt,
    messages: [{role:"user", content: userPrompt}]
  };
  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var resp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code !== 200) {
      var errMsg = "Anthropic API error (" + code + ")";
      try { var errJson = JSON.parse(body); if (errJson.error && errJson.error.message) errMsg += ": " + errJson.error.message; } catch(e2) {}
      return {ok:false, error: errMsg};
    }
    var data = JSON.parse(body);
    var textBlock = "";
    for (var i=0;i<(data.content||[]).length;i++) {
      if (data.content[i].type === "text") { textBlock += data.content[i].text; }
    }
    var cleaned = textBlock.trim().replace(/^```json/i,"").replace(/^```/,"").replace(/```$/,"").trim();
    var parsed;
    try { parsed = JSON.parse(cleaned); } catch(e3) {
      return {ok:false, error:"AI response could not be parsed as JSON. Try rephrasing your request."};
    }
    return {ok:true, items: parsed.items||[], explanation: parsed.explanation||""};
  } catch(err) {
    return {ok:false, error:"AI request failed: "+err.toString()};
  }
}

// ── BOOTSTRAP — combines everything the frontend needs on app load into a
// single round trip (SOR versions, Non-SOR items, T-Codes, Processes,
// Estimates, charge rates, and — for admins — Users). Cuts boot latency
// from ~6-7 sequential/parallel Apps Script round trips down to one,
// since each individual getX() call still pays its own Apps Script
// execution overhead even when fired in parallel from the client.
function bootstrap(userId, role) {
  var sorV = getSORVersions();
  var nonSor = getNonSOR();
  var tcodes = getTCodes();
  var procs = getProcesses();
  var estimates = getEstimates(userId, role);
  var rates = getChargeRates();
  var result = {
    ok: true,
    sorVersions: sorV.ok ? sorV.versions : [],
    sorVersionCounts: sorV.ok ? sorV.counts : {},
    nonSor: nonSor.ok ? nonSor.items : [],
    tcodes: tcodes.ok ? tcodes.items : [],
    processes: procs.ok ? procs.items : [],
    estimates: estimates.ok ? estimates.estimates : [],
    chargeRates: rates.ok ? rates.rates : defaultChargeRates(),
    aiAllowed: isUserAiAllowed_(userId, role),
    aiKeyConfigured: !!getConfigValue_("anthropicApiKey", "")
  };
  if (role === "admin") {
    var users = getUsers();
    result.users = users.ok ? users.users : [];
  }
  return result;
}

function getChargeRates() {
  var ss = getOrCreateSS();
  var sheet = getConfigSheet_(ss);
  var data = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues() : [];
  for (var i=0;i<data.length;i++) {
    if (data[i][0] === "chargeRates") return {ok:true, rates: tryParse(data[i][1], defaultChargeRates())};
  }
  sheet.appendRow(["chargeRates", JSON.stringify(defaultChargeRates())]);
  return {ok:true, rates: defaultChargeRates()};
}
function saveChargeRates(rates, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can change system charge rates"};
  var ss = getOrCreateSS();
  var sheet = getConfigSheet_(ss);
  var data = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues() : [];
  for (var i=0;i<data.length;i++) {
    if (data[i][0] === "chargeRates") {
      sheet.getRange(i+2,2).setValue(JSON.stringify(rates));
      return {ok:true};
    }
  }
  sheet.appendRow(["chargeRates", JSON.stringify(rates)]);
  return {ok:true};
}
function resetChargeRates() {
  var ss = getOrCreateSS();
  var sheet = getConfigSheet_(ss);
  var data = sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues() : [];
  var def = defaultChargeRates();
  for (var i=0;i<data.length;i++) {
    if (data[i][0] === "chargeRates") {
      sheet.getRange(i+2,2).setValue(JSON.stringify(def));
      return {ok:true, rates: def};
    }
  }
  sheet.appendRow(["chargeRates", JSON.stringify(def)]);
  return {ok:true, rates: def};
}

// ── SOR ────────────────────────────────────────────────────────
function getSOR(filterVersion) {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_SOR);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, items:[], total:0};
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, SOR_HDR.length).getValues();
  var items = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!r[0] || String(r[0]).trim() === "") continue;
    if (filterVersion && String(r[5]).trim() !== filterVersion) continue;
    items.push({
      item_no:     String(r[0]).trim(),
      description: String(r[1]).trim(),
      uom:         String(r[2]).trim(),
      rate:        parseFloat(r[3]) || 0,
      section:     String(r[4]).trim(),
      version:     String(r[5]).trim()
    });
  }
  return {ok:true, items:items, total:items.length};
}

function getSORVersions() {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_SOR);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, versions:[]};
  var data = sheet.getRange(2, 6, sheet.getLastRow()-1, 1).getValues();
  var seen = {}, counts = {}, versions = [];
  for (var i = 0; i < data.length; i++) {
    var v = String(data[i][0]).trim();
    if (!v) continue;
    if (!seen[v]) { seen[v]=true; versions.push(v); counts[v]=0; }
    counts[v]++;
  }
  versions.sort(); // lexicographic; caller resolves "latest" via [...].sort().pop()
  return {ok:true, versions:versions, counts:counts};
}

function importSOR(items, version) {
  if (!items || !items.length) return {ok:false, error:"No items"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_SOR, SOR_HDR);
  var existing = {};
  if (sheet.getLastRow() > 1) {
    var ex = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    for (var i = 0; i < ex.length; i++) {
      existing[String(ex[i][0]).trim()+"__"+String(ex[i][5]).trim()] = true;
    }
  }
  var toWrite = [];
  var versionsSeen = {};
  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    // Per-row version (from an uploaded "version" column) takes priority;
    // falls back to the single version name supplied for the whole batch.
    var itemVersion = (item.version && String(item.version).trim()) || version;
    var key  = String(item.item_no).trim()+"__"+String(itemVersion).trim();
    if (!existing[key]) {
      toWrite.push([item.item_no, item.description, item.uom,
                    item.rate, item.section, itemVersion]);
      existing[key] = true;
    }
    versionsSeen[itemVersion] = (versionsSeen[itemVersion]||0) + 1;
  }
  if (toWrite.length > 0) {
    var bSize = 500;
    for (var b = 0; b < toWrite.length; b += bSize) {
      var batch = toWrite.slice(b, b+bSize);
      sheet.getRange(sheet.getLastRow()+1, 1, batch.length, SOR_HDR.length)
           .setValues(batch);
    }
  }
  return {
    ok:true, added: toWrite.length,
    skipped: items.length - toWrite.length,
    total:   sheet.getLastRow()-1,
    versions: Object.keys(versionsSeen)
  };
}

// ── NON-SOR ITEMS ────────────────────────────────────────────────
function getNonSOR() {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_NONSOR);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, items:[]};
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,NONSOR_HDR.length).getValues();
  var items = [];
  for (var i=0;i<data.length;i++) {
    var r = data[i];
    if (!r[0]) continue;
    items.push({
      item_no:     String(r[0]).trim(),
      description: String(r[1]).trim(),
      uom:         String(r[2]).trim(),
      rate:        parseFloat(r[3]) || 0,
      section:     String(r[4]).trim(),
      remarks:     String(r[5]||"").trim(),
      addedBy:     String(r[6]||"").trim(),
      addedAt:     String(r[7]||"").trim(),
      version:     "NONSOR"
    });
  }
  return {ok:true, items:items};
}

function findNonSORRow_(sheet, item_no) {
  if (sheet.getLastRow() < 2) return -1;
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  for (var i=0;i<data.length;i++) {
    if (String(data[i][0]).trim() === String(item_no).trim()) return i+2;
  }
  return -1;
}

function addNonSOR(item) {
  if (!item || !item.item_no || !item.description) return {ok:false, error:"Missing fields"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_NONSOR, NONSOR_HDR);
  if (findNonSORRow_(sheet, item.item_no) > -1) {
    return {ok:false, error:"Item number already exists"};
  }
  sheet.appendRow([
    item.item_no, item.description, item.uom||"", item.rate||0,
    item.section||"", item.remarks||"", item.addedBy||"", new Date().toISOString()
  ]);
  return {ok:true};
}

function updateNonSOR(item, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can edit Non-SOR items"};
  if (!item || !item.item_no) return {ok:false, error:"Missing item_no"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_NONSOR, NONSOR_HDR);
  var origNo = item.origItemNo || item.item_no;
  var row = findNonSORRow_(sheet, origNo);
  if (row === -1) return {ok:false, error:"Item not found"};
  sheet.getRange(row,1,1,8).setValues([[
    item.item_no, item.description, item.uom||"", item.rate||0,
    item.section||"", item.remarks||"",
    sheet.getRange(row,7).getValue(), sheet.getRange(row,8).getValue()
  ]]);
  return {ok:true};
}

function deleteNonSOR(item_no, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can delete Non-SOR items"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_NONSOR, NONSOR_HDR);
  var row = findNonSORRow_(sheet, item_no);
  if (row === -1) return {ok:false, error:"Item not found"};
  sheet.deleteRow(row);
  return {ok:true};
}

// ── T-CODES ────────────────────────────────────────────────────
function getTCodes() {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_TCODES);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, items:[]};
  var n = Math.max(TCODE_HDR.length, sheet.getLastColumn());
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,n).getValues();
  var items = [];
  for (var i=0;i<data.length;i++) {
    var r = data[i];
    if (!r[0]) continue;
    items.push({
      tcode: String(r[0]).trim(), purpose: String(r[1]||"").trim(),
      description: String(r[2]||"").trim(), remarks: String(r[3]||"").trim(),
      addedBy: String(r[4]||"").trim(), addedAt: String(r[5]||"").trim(),
      image: String(r[7]||"").trim()
    });
  }
  return {ok:true, items:items};
}
function findTCodeRow_(sheet, tcode) {
  if (sheet.getLastRow() < 2) return -1;
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  for (var i=0;i<data.length;i++) if (String(data[i][0]).trim()===String(tcode).trim()) return i+2;
  return -1;
}
function addTCode(item) {
  if (!item || !item.tcode) return {ok:false, error:"Missing T Code"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_TCODES, TCODE_HDR);
  if (findTCodeRow_(sheet, item.tcode) > -1) return {ok:false, error:"T Code already exists"};
  sheet.appendRow([item.tcode, item.purpose||"", item.description||"",
                    item.remarks||"", item.addedBy||"", new Date().toISOString(), "approved", item.image||""]);
  return {ok:true};
}
function canEditEntry_(rowOwner, userId, role) {
  return role === "admin" || rowOwner === userId;
}
function updateTCode(item, userId, role) {
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_TCODES, TCODE_HDR);
  var row = findTCodeRow_(sheet, item.origTcode || item.tcode);
  if (row === -1) return {ok:false, error:"Not found"};
  var owner = String(sheet.getRange(row,5).getValue());
  if (!canEditEntry_(owner, userId, role)) return {ok:false, error:"Not permitted"};
  var curImage = item.image !== undefined ? item.image : String(sheet.getRange(row,8).getValue()||"");
  sheet.getRange(row,1,1,8).setValues([[item.tcode, item.purpose||"", item.description||"",
                    item.remarks||"", owner, sheet.getRange(row,6).getValue(), "approved", curImage]]);
  return {ok:true};
}
function deleteTCode(tcode, userId, role) {
  if (role !== "admin") return {ok:false, error:"Only admin can delete"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_TCODES, TCODE_HDR);
  var row = findTCodeRow_(sheet, tcode);
  if (row === -1) return {ok:false, error:"Not found"};
  sheet.deleteRow(row);
  return {ok:true};
}

// ── PROCESSES ──────────────────────────────────────────────────
function getProcesses() {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_PROCESSES);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, items:[]};
  var n = Math.max(PROC_HDR.length, sheet.getLastColumn());
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,n).getValues();
  var items = [];
  for (var i=0;i<data.length;i++) {
    var r = data[i];
    if (!r[0]) continue;
    items.push({
      name: String(r[0]).trim(), description: String(r[1]||"").trim(),
      remarks: String(r[2]||"").trim(), addedBy: String(r[3]||"").trim(),
      addedAt: String(r[4]||"").trim(),
      image: String(r[6]||"").trim()
    });
  }
  return {ok:true, items:items};
}
function findProcessRow_(sheet, name) {
  if (sheet.getLastRow() < 2) return -1;
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  for (var i=0;i<data.length;i++) if (String(data[i][0]).trim()===String(name).trim()) return i+2;
  return -1;
}
function addProcess(item) {
  if (!item || !item.name) return {ok:false, error:"Missing name"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_PROCESSES, PROC_HDR);
  if (findProcessRow_(sheet, item.name) > -1) return {ok:false, error:"Process already exists"};
  sheet.appendRow([item.name, item.description||"", item.remarks||"",
                    item.addedBy||"", new Date().toISOString(), "approved", item.image||""]);
  return {ok:true};
}
function updateProcess(item, userId, role) {
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_PROCESSES, PROC_HDR);
  var row = findProcessRow_(sheet, item.origName || item.name);
  if (row === -1) return {ok:false, error:"Not found"};
  var owner = String(sheet.getRange(row,4).getValue());
  if (!canEditEntry_(owner, userId, role)) return {ok:false, error:"Not permitted"};
  var curImage = item.image !== undefined ? item.image : String(sheet.getRange(row,7).getValue()||"");
  sheet.getRange(row,1,1,7).setValues([[item.name, item.description||"", item.remarks||"",
                    owner, sheet.getRange(row,5).getValue(), "approved", curImage]]);
  return {ok:true};
}
function deleteProcess(name, userId, role) {
  if (role !== "admin") return {ok:false, error:"Only admin can delete"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_PROCESSES, PROC_HDR);
  var row = findProcessRow_(sheet, name);
  if (row === -1) return {ok:false, error:"Not found"};
  sheet.deleteRow(row);
  return {ok:true};
}

// ── USERS ──────────────────────────────────────────────────────
function getUsers() {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, users:[]};
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,USER_HDR.length).getValues();
  var users = [];
  for (var i=0;i<data.length;i++) {
    var r = data[i];
    if (!r[0]) continue;
    users.push({id:String(r[0]).trim(), name:String(r[1]).trim(), role:String(r[3]).trim(),
                status:String(r[4]).trim(), createdAt:String(r[5]).trim()});
  }
  return {ok:true, users:users};
}
function findUserRow_(sheet, id) {
  if (sheet.getLastRow() < 2) return -1;
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  for (var i=0;i<data.length;i++) if (String(data[i][0]).trim()===String(id).trim()) return i+2;
  return -1;
}
function registerUser(user) {
  if (!user || !user.id || !user.passHash) return {ok:false, error:"Missing fields"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_USERS, USER_HDR);
  ensureSeedAdmin_(sheet);
  if (findUserRow_(sheet, user.id) > -1) return {ok:false, error:"User ID already exists"};
  sheet.appendRow([user.id, user.name||user.id, user.passHash, "user", "pending", new Date().toISOString()]);
  return {ok:true, message:"Registered. Awaiting admin approval."};
}
function loginUser(id, passHash) {
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_USERS, USER_HDR);
  ensureSeedAdmin_(sheet);
  var row = findUserRow_(sheet, id);
  if (row === -1) return {ok:false, error:"User not found"};
  var r = sheet.getRange(row,1,1,USER_HDR.length).getValues()[0];
  if (String(r[2]) !== String(passHash)) return {ok:false, error:"Incorrect password"};
  if (String(r[4]) === "pending") return {ok:false, error:"Account pending admin approval"};
  if (String(r[4]) === "rejected") return {ok:false, error:"Account access denied"};
  return {ok:true, user:{id:String(r[0]), name:String(r[1]), role:String(r[3]), status:String(r[4])}};
}
// Self-healing seed: if the Users sheet has no data rows at all (fresh
// deployment where /exec?action=init was never explicitly called), seed the
// default admin so the very first login never dead-ends. Only acts when the
// sheet is completely empty of data rows — never touches an existing one.
function ensureSeedAdmin_(userSheet) {
  if (userSheet.getLastRow() < 2) {
    userSheet.appendRow(["admin","Administrator",hashPassGS("admin123"),"admin","approved",new Date().toISOString()]);
  }
}
function changePassword(id, oldHash, newHash) {
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_USERS, USER_HDR);
  var row = findUserRow_(sheet, id);
  if (row === -1) return {ok:false, error:"User not found"};
  var cur = sheet.getRange(row,3).getValue();
  if (String(cur) !== String(oldHash)) return {ok:false, error:"Current password incorrect"};
  sheet.getRange(row,3).setValue(newHash);
  return {ok:true};
}
function saveUser(user, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can edit users"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_USERS, USER_HDR);
  var row = findUserRow_(sheet, user.id);
  if (row === -1) return {ok:false, error:"User not found"};
  if (user.name)     sheet.getRange(row,2).setValue(user.name);
  if (user.role)      sheet.getRange(row,4).setValue(user.role);
  if (user.passHash)  sheet.getRange(row,3).setValue(user.passHash);
  if (user.status)    sheet.getRange(row,5).setValue(user.status);
  return {ok:true};
}
function approveUser(id, approve, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can approve users"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_USERS, USER_HDR);
  var row = findUserRow_(sheet, id);
  if (row === -1) return {ok:false, error:"User not found"};
  sheet.getRange(row,5).setValue(approve ? "approved" : "rejected");
  return {ok:true};
}
function deleteUser(id, actingRole) {
  if (actingRole !== "admin") return {ok:false, error:"Only admin can delete users"};
  if (id === "admin") return {ok:false, error:"Cannot delete the primary admin account"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_USERS, USER_HDR);
  var row = findUserRow_(sheet, id);
  if (row === -1) return {ok:false, error:"User not found"};
  sheet.deleteRow(row);
  return {ok:true};
}

// ── ESTIMATES ──────────────────────────────────────────────────
function rowToEstimateSummary_(r) {
  return {
    id:String(r[0]), name:String(r[1]), number:String(r[2]), date:String(r[3]),
    division:String(r[4]), ref:String(r[5]), work:String(r[6]),
    sorVersions: tryParse(r[7], []), addCharge: tryParse(r[8], {}),
    settings: tryParse(r[9], {}), finalTotal: parseFloat(r[10])||0,
    savedAt:String(r[11]), status:String(r[12])||"draft", itemCount: parseInt(r[13])||0,
    ownerId:String(r[14]), ownerName:String(r[15]),
    audited: r[16]===true||r[16]==="true",
    remark1:String(r[19]||""), remark2:String(r[20]||""), remark3:String(r[21]||""),
    remark4:String(r[22]||""), remark5:String(r[23]||""),
    viewerVisible: r[24]===true||r[24]==="true",
    estimateType: String(r[25]||"project"),
    amcAmount: parseFloat(r[26])||0,
    progressImages: tryParse(r[27], []),
    supplySpares: r[28]===true||r[28]==="true",
    amcRemark: String(r[29]||""),
    progress: tryParse(r[30], defaultProgress_()),
    supplyTotal: parseFloat(r[31])||0, projectTotal: parseFloat(r[32])||0,
    omTotal: parseFloat(r[33])||0, nonSorSupplyTotal: parseFloat(r[34])||0,
    nonSorOmTotal: parseFloat(r[35])||0,
    trLineMode: r[36]===true||r[36]==="true",
    landCompChargePct: r[37]!==undefined && r[37]!=="" ? parseFloat(r[37]) : 15,
    centagesOverridden: r[38]===true||r[38]==="true",
    pmoPrEntries: tryParse(r[39], []),
    prSettings: tryParse(r[40], null)
  };
}
function defaultProgress_() {
  var stages = {};
  for (var i=1;i<=25;i++) stages[i] = {done:false, remarks:"", link:""};
  return {
    stages:stages, woAmount:0, bills:[], orderClosed:false, orderClosedRemarks:"",
    orderDetails: defaultOrderDetails_()
  };
}
function defaultOrderDetails_() {
  return {
    adminApprovalNo:"", adminApprovalDate:"", expenseBookingFY:"",
    processedVia:"", processedViaOther:"",
    tenderNo:"",
    particularsOfWork:"", msetclEstimatedCost:"", rfxNo:"", nameOfOffice:"",
    noOfBidsReceived:"", l1Rates:"", vendorName:"", vendorCode:"",
    loiNo:"", loiDate:"",
    woNo:"", woDate:"", poNo:"", poDate:"", woRemarks:"",
    stampDutyPaid:"", stampNo:"", stampDate:"", stampRemarks:"",
    sdBgNo:"", nameOfBank:"", sdBgOtherDetails:"",
    preparedBy:""
  };
}
function getEstimates(userId, role) {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_ESTIMATES);
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, estimates:[]};
  var n = Math.max(EST_HDR.length, sheet.getLastColumn());
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,n).getValues();
  var out = [];
  for (var i=0;i<data.length;i++) {
    var r = data[i];
    if (!r[0]) continue;
    var est = rowToEstimateSummary_(r);
    if (role === "admin") out.push(est);
    else if (role === "viewer") { if (est.viewerVisible) out.push(est); }
    else { if (est.ownerId === userId) out.push(est); }
  }
  return {ok:true, estimates:out};
}
function getEstimate(id, userId, role) {
  var ss = getOrCreateSS();
  var sheet = ss.getSheetByName(SHEET_ESTIMATES);
  if (!sheet || sheet.getLastRow() < 2) return {ok:false, error:"Not found"};
  var n = Math.max(EST_HDR.length, sheet.getLastColumn());
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,n).getValues();
  for (var i=0;i<data.length;i++) {
    if (String(data[i][0]) === String(id)) {
      var est = rowToEstimateSummary_(data[i]);
      if (role !== "admin" && est.ownerId !== userId && !(role==="viewer" && est.viewerVisible)) {
        return {ok:false, error:"Access denied"};
      }
      var itemSheet = ss.getSheetByName(SHEET_ITEMS);
      var items = [];
      if (itemSheet && itemSheet.getLastRow() > 1) {
        var idata = itemSheet.getRange(2,1,itemSheet.getLastRow()-1,ITEM_HDR.length).getValues();
        for (var j=0;j<idata.length;j++) {
          if (String(idata[j][0]) === String(id)) {
            items.push({
              item_no:String(idata[j][1]), version:String(idata[j][2]),
              description:String(idata[j][3]), uom:String(idata[j][4]),
              rate:parseFloat(idata[j][5])||0, qty:parseFloat(idata[j][6])||0,
              section:String(idata[j][7]), remarks:String(idata[j][8]||""),
              isCustom: idata[j][9]===true||idata[j][9]==="true"
            });
          }
        }
      }
      return {ok:true, estimate:est, items:items};
    }
  }
  return {ok:false, error:"Not found"};
}
function saveEstimate(estimate, items) {
  if (!estimate) return {ok:false, error:"Missing estimate"};
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_ESTIMATES, EST_HDR);
  var isNew = !estimate.id;
  var id = estimate.id || ("EST"+new Date().getTime());
  var row = isNew ? -1 : findEstimateRow_(sheet, id);
  var progress = estimate.progress || defaultProgress_();
  var rowData = [
    id, estimate.name||"", estimate.number||"", estimate.date||"",
    estimate.division||"", estimate.ref||"", estimate.work||"",
    JSON.stringify(estimate.sorVersions||[]), JSON.stringify(estimate.addCharge||{}),
    JSON.stringify(estimate.settings||{}), estimate.finalTotal||0,
    new Date().toISOString(), estimate.status||"draft", (items||[]).length,
    estimate.ownerId||"", estimate.ownerName||"",
    estimate.audited?true:false, "", "",
    estimate.remark1||"", estimate.remark2||"", estimate.remark3||"",
    estimate.remark4||"", estimate.remark5||"",
    estimate.viewerVisible?true:false, estimate.estimateType||"project",
    estimate.amcAmount||0, JSON.stringify(estimate.progressImages||[]),
    estimate.supplySpares?true:false, estimate.amcRemark||"",
    JSON.stringify(progress),
    estimate.supplyTotal||0, estimate.projectTotal||0, estimate.omTotal||0,
    estimate.nonSorSupplyTotal||0, estimate.nonSorOmTotal||0,
    estimate.trLineMode?true:false, estimate.landCompChargePct||15,
    estimate.centagesOverridden?true:false,
    JSON.stringify(estimate.pmoPrEntries||[]),
    JSON.stringify(estimate.prSettings||null)
  ];
  if (row === -1) {
    sheet.appendRow(rowData);
  } else {
    sheet.getRange(row,1,1,rowData.length).setValues([rowData]);
  }
  // Replace item rows for this estimate
  var itemSheet = getOrCreateSheet(ss, SHEET_ITEMS, ITEM_HDR);
  if (itemSheet.getLastRow() > 1) {
    var idata = itemSheet.getRange(2,1,itemSheet.getLastRow()-1,1).getValues();
    var rowsToDelete = [];
    for (var i=0;i<idata.length;i++) if (String(idata[i][0])===String(id)) rowsToDelete.push(i+2);
    for (var k=rowsToDelete.length-1;k>=0;k--) itemSheet.deleteRow(rowsToDelete[k]);
  }
  if (items && items.length) {
    var toWrite = items.map(function(it){
      return [id, it.item_no, it.version||"", it.description, it.uom,
              it.rate, it.qty, it.section, it.remarks||"", it.isCustom?true:false];
    });
    itemSheet.getRange(itemSheet.getLastRow()+1,1,toWrite.length,ITEM_HDR.length).setValues(toWrite);
  }
  return {ok:true, id:id};
}
function findEstimateRow_(sheet, id) {
  if (sheet.getLastRow() < 2) return -1;
  var data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  for (var i=0;i<data.length;i++) if (String(data[i][0])===String(id)) return i+2;
  return -1;
}
function deleteEstimate(id, userId, role) {
  var ss = getOrCreateSS();
  var sheet = getOrCreateSheet(ss, SHEET_ESTIMATES, EST_HDR);
  var row = findEstimateRow_(sheet, id);
  if (row === -1) return {ok:false, error:"Not found"};
  var owner = String(sheet.getRange(row,15).getValue());
  if (role !== "admin" && owner !== userId) return {ok:false, error:"Not permitted"};
  sheet.deleteRow(row);
  var itemSheet = ss.getSheetByName(SHEET_ITEMS);
  if (itemSheet && itemSheet.getLastRow() > 1) {
    var idata = itemSheet.getRange(2,1,itemSheet.getLastRow()-1,1).getValues();
    var rowsToDelete = [];
    for (var i=0;i<idata.length;i++) if (String(idata[i][0])===String(id)) rowsToDelete.push(i+2);
    for (var k=rowsToDelete.length-1;k>=0;k--) itemSheet.deleteRow(rowsToDelete[k]);
  }
  return {ok:true};
}
