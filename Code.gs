/**
 * Shopping List - Backend
 * All operations via doGet for proper CORS handling.
 */

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const RECIPIENT_EMAIL = PropertiesService.getScriptProperties().getProperty("RECIPIENT_EMAIL") || "jsknight007@gmail.com";
const TIMEZONE = Session.getScriptTimeZone();
const STORES = ["Hy-Vee", "Fareway", "Target", "Wal-Mart", "Costco"];

function getSheet() {
  if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID not set in Script Properties");
  return SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || "fetch";

    switch (action) {
      case "add": return handleAdd(params);
      case "update": return handleUpdate(params);
      case "delete": return handleDelete(params);
      case "email_demand": return handleEmail(params);
      default: return handleFetch();
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function handleAdd(params) {
  var text = params.text;
  if (!text || text.trim() === "") {
    return jsonResponse({ error: "Item text is required" });
  }
  text = text.trim();
  if (text.length > 200) {
    return jsonResponse({ error: "Item too long (max 200 chars)" });
  }

  var parts = text.split(" ");
  var name = parts[0];
  var item = parts.slice(1).join(" ");
  if (!item) { item = name; name = "General"; }

  var sheet = getSheet();
  sheet.appendRow([new Date(), name, item, ""]);
  return jsonResponse({ status: "success", name: name, item: item });
}

function handleUpdate(params) {
  var index = parseInt(params.index);
  var value = params.value || "";

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();

  if (isNaN(index) || index < 2 || index > lastRow) {
    return jsonResponse({ error: "Invalid row index" });
  }
  if (value && STORES.indexOf(value) === -1) {
    return jsonResponse({ error: "Invalid store value" });
  }

  sheet.getRange(index, 4).setValue(value);
  return jsonResponse({ status: "success" });
}

function handleDelete(params) {
  var index = parseInt(params.index);
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();

  if (isNaN(index) || index < 2 || index > lastRow) {
    return jsonResponse({ error: "Invalid row index" });
  }

  if (params.name && params.item) {
    var row = sheet.getRange(index, 1, 1, 4).getValues()[0];
    var rowName = (row[1] || "").toString().toLowerCase().trim();
    var rowItem = (row[2] || "").toString().toLowerCase().trim();
    if (rowName !== params.name.toLowerCase().trim() || rowItem !== params.item.toLowerCase().trim()) {
      return jsonResponse({ error: "Row content mismatch - list may have changed" });
    }
  }

  sheet.deleteRow(index);
  return jsonResponse({ status: "success" });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ error: "Invalid request" });
    }
    var data = JSON.parse(e.postData.contents);
    if (data.action === "email_demand") {
      return handleEmail(data);
    }
    return jsonResponse({ error: "Unknown POST action" });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function handleEmail(params) {
  var store = params.store || "All";
  var items = params.items || "";
  if (!items) {
    return jsonResponse({ error: "No items to email" });
  }
  if (items.length > 5000) {
    return jsonResponse({ error: "Item list too long" });
  }

  var body = "Items for " + store + ":\n\n" + items;
  MailApp.sendEmail(RECIPIENT_EMAIL, "Shopping List: " + store, body);
  return jsonResponse({ status: "success" });
}

function handleFetch() {
  var sheet = getSheet();
  var fullData = sheet.getDataRange().getValues();

  if (fullData.length <= 1) {
    return jsonResponse([]);
  }

  var rows = fullData.slice(1);
  var list = rows.map(function(row, index) {
    var dateVal = row[0];
    var formattedDate = "N/A";

    try {
      if (dateVal instanceof Date) {
        formattedDate = Utilities.formatDate(dateVal, TIMEZONE, "MM/dd");
      } else if (dateVal && dateVal !== "") {
        var convertedDate = new Date(dateVal);
        if (!isNaN(convertedDate.getTime())) {
          formattedDate = Utilities.formatDate(convertedDate, TIMEZONE, "MM/dd");
        }
      }
    } catch (e) {
      formattedDate = "Err";
    }

    return {
      id: index + 2,
      date: formattedDate,
      name: row[1] || "",
      item: row[2] || "",
      store: row[3] || ""
    };
  });

  return jsonResponse(list);
}

/**
 * Trigger-based: Groups by store and emails summary.
 * Set up in Triggers (Clock icon).
 */
function sendDailySummary() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var rows = data.slice(1);
  var stores = {};

  rows.forEach(function(row) {
    var store = row[3] || "Unassigned/New";
    if (!stores[store]) stores[store] = [];
    stores[store].push(row[1] + ": " + row[2]);
  });

  var emailBody = "<div style='font-family:sans-serif;'><h2>Today's Needs List</h2>";
  for (var store in stores) {
    emailBody += "<h3 style='color:#2b7de9;'>" + store + "</h3><ul>";
    stores[store].forEach(function(item) {
      emailBody += "<li style='margin-bottom:5px;'>" + escapeHtml(item) + "</li>";
    });
    emailBody += "</ul>";
  }
  emailBody += "</div>";

  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: "Grocery & Needs Summary",
    htmlBody: emailBody
  });
}

function escapeHtml(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
