/**
 * AyurVani lead capture — Google Apps Script Web App endpoint.
 * Deploy: Deploy > New deployment > Web app > Execute as "Me",
 * Who has access "Anyone". Paste the /exec URL into GOOGLE_SCRIPT_URL in index.html.
 *
 * Field keys below must match the <input name="..."> values in the page's #orderForm.
 */

var FIELDS = [
  { key: "name",            header: "Name" },
  { key: "mobile",          header: "Mobile Number" },   // was "moblie" — typo meant this column never filled
  { key: "product",         header: "Product" },         // page posts this
  { key: "source",          header: "Source" },
  { key: "utm_source",      header: "UTM Source" },
  { key: "utm_medium",      header: "UTM Medium" },
  { key: "utm_campaign",    header: "UTM Campaign" },
  { key: "utm_campaign_id", header: "UTM Campaign ID" },
  { key: "utm_adset",       header: "UTM Ad Set" },
  { key: "utm_adset_id",    header: "UTM Ad Set ID" },
  { key: "utm_ad",          header: "UTM Ad" },
  { key: "utm_ad_id",       header: "UTM Ad ID" },
  { key: "utm_placement",   header: "UTM Placement" },
  { key: "fbclid",          header: "fbclid" },          // page posts these three for Meta attribution
  { key: "fbc",             header: "fbc" },
  { key: "fbp",             header: "fbp" },
  { key: "page_url",        header: "Page URL" },
  { key: "user_agent",      header: "User Agent" },
];

/* Letters only — Latin and Devanagari, single spaces between words. The page is
   Hindi-first, so Devanagari must be allowed or no Hindi name passes; the range
   skips U+0964-U+096F, which is the danda and the Devanagari digits. */
var NAME_RE   = /^[A-Za-zऀ-ॣ॰-ॿ]+(?: [A-Za-zऀ-ॣ॰-ॿ]+)*$/;
var MOBILE_RE = /^[6-9][0-9]{9}$/;

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Meta traffic can fire several leads in the same second — serialise writes
    // so two submissions never land on the same row.
    lock.waitLock(20000);

    var params = readParams_(e);

    // Validate again on the server. The browser checks are for the person
    // filling the form; these are for anything that posts straight to the URL.
    var name = String(params.name || "").trim().replace(/\s+/g, " ");
    if (name.length < 2 || name.length > 60 || !NAME_RE.test(name)) {
      return jsonResponse_({ result: "error", message: "Invalid name." });
    }

    // The form field is called "mobile"; "phone" kept as a fallback.
    var mobile = normaliseMobile_(params.mobile || params.phone);
    if (!MOBILE_RE.test(mobile)) {
      return jsonResponse_({ result: "error", message: "Invalid mobile number." });
    }

    params.name = name;
    params.mobile = mobile;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ensureHeaderRow_(sheet);

    var row = [new Date()];
    FIELDS.forEach(function (field) {
      row.push(params[field.key] || "");
    });
    sheet.appendRow(row);

    return jsonResponse_({ result: "success" });
  } catch (err) {
    return jsonResponse_({ result: "error", message: err.message });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function doGet() {
  return jsonResponse_({ result: "error", message: "This endpoint only accepts POST requests." });
}

/** Accepts form-encoded posts (e.parameter), raw JSON bodies, and multipart FormData. */
function readParams_(e) {
  if (e && e.postData && e.postData.type === "application/json") {
    try { return JSON.parse(e.postData.contents) || {}; } catch (err) { /* fall through */ }
  }
  if (e && e.parameter && Object.keys(e.parameter).length) return e.parameter;
  // fetch(..., { body: new FormData() }) sends multipart/form-data, which Apps
  // Script does not populate e.parameter from — parse the raw body instead.
  if (e && e.postData && String(e.postData.type || "").indexOf("multipart/form-data") === 0) {
    return parseMultipart_(e.postData.contents, e.postData.type);
  }
  return (e && e.parameter) || {};
}

function parseMultipart_(body, contentType) {
  var out = {};
  var match = /boundary=(?:"([^"]+)"|([^;]+))/.exec(contentType || "");
  if (!match || !body) return out;
  var boundary = "--" + (match[1] || match[2]).trim();
  body.split(boundary).forEach(function (part) {
    var nameMatch = /name="([^"]*)"/.exec(part);
    if (!nameMatch) return;
    var split = part.indexOf("\r\n\r\n");
    if (split === -1) return;
    out[nameMatch[1]] = part.slice(split + 4).replace(/\r\n$/, "");
  });
  return out;
}

/** Strips +91 / 0 prefixes, spaces and dashes down to the bare 10 digits. */
function normaliseMobile_(value) {
  var digits = String(value || "").replace(/\D/g, "");
  if (digits.length === 12 && digits.indexOf("91") === 0) { digits = digits.slice(2); }
  if (digits.length === 11 && digits.charAt(0) === "0")   { digits = digits.slice(1); }
  return digits;
}

function ensureHeaderRow_(sheet) {
  if (sheet.getLastRow() > 0) return;
  var headers = ["Timestamp"].concat(FIELDS.map(function (field) { return field.header; }));
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
