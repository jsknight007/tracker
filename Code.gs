function doPost(e) {
  const command = e.parameter.command;
  const result = parseCommand(command);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  
  // COLUMN MAPPING (A through F)
  sheet.appendRow([
    new Date(),      // A: Entry Time
    result.date,     // B: When (Transaction Date)
    result.store,    // C: Where
    result.price,    // D: How Much
    result.card,     // E: Which Card
    command          // F: Raw Command
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({status: "success"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const action = e.parameter.action;

  if (action === "deleteLast") {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRow(lastRow);
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "deleteSpecific") {
    const index = parseInt(e.parameter.index);
    sheet.deleteRow(index);
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }

  const cardFilter = e.parameter.card;
  const data = sheet.getDataRange().getValues();
  const entries = [];

  for (let i = 1; i < data.length; i++) {
    let rowCard = data[i][4] ? data[i][4].toString().trim().toLowerCase() : "";
    let filterCard = cardFilter ? cardFilter.toString().trim().toLowerCase() : "";

    if (rowCard === filterCard) {
      entries.push({
        index: i + 1,
        when: Utilities.formatDate(new Date(data[i][1]), "GMT-6", "MM/dd/yy"),
        where: data[i][2],
        price: data[i][3]
      });
    }
  }

  return ContentService.createTextOutput(JSON.stringify({entries: entries.reverse().slice(0, 10)}))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseCommand(command) {
  const result = {
    price: 0.00,
    store: "Unknown Store",
    date: new Date(),
    card: "Other",
    success: false
  };

  if (!command) return result;
  const lower = command.toLowerCase().trim();

  // 1. PRICE ANCHOR
  const priceMatch = command.match(/\$?\s?(\d+(\.\d{2})?)/);
  if (priceMatch) {
    result.price = parseFloat(priceMatch[1]);
    result.success = true;
    let storePart = command.substring(0, priceMatch.index).trim();
    result.store = storePart.replace(/^(at|to|from|for)\s+/i, "");
  }

  // 2. ROBUST DATE LOGIC (Fixes the "One Day Early" bug)
  const dateRegex = /(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+(\d{1,2})(st|nd|rd|th)?/i;
  const dateFound = command.match(dateRegex);
  
  if (dateFound) {
    const monthStr = dateFound[1].toLowerCase();
    const day = parseInt(dateFound[2]);
    const year = new Date().getFullYear();
    
    const monthMap = {
      jan:0, january:0, feb:1, february:1, mar:2, march:2, apr:3, april:3, 
      may:4, jun:5, june:5, jul:6, july:6, aug:7, august:7, sep:8, september:8, 
      oct:9, october:9, nov:10, november:10, dec:11, december:11
    };

    // Create date relative to script's timezone, set to Noon to avoid UTC drift
    const localDate = new Date();
    localDate.setFullYear(year);
    localDate.setMonth(monthMap[monthStr]);
    localDate.setDate(day);
    localDate.setHours(12, 0, 0, 0); 
    
    result.date = localDate;
  } else if (lower.includes("yesterday")) {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(12, 0, 0, 0);
    result.date = yesterday;
  }

  // 3. STORE CLEANUP
  const storeMap = {
    "panera": "Panera Bread", "target": "Target", "costco": "Costco",
    "hyvee": "Hy-Vee", "hy-vee": "Hy-Vee", "amazon": "Amazon",
    "walmart": "Walmart", "granite city": "Granite City"
  };

  let rawStore = result.store.toLowerCase().trim();
  let lookupKey = rawStore.replace(/[\s-]/g, ""); 
  
  if (storeMap[rawStore]) {
    result.store = storeMap[rawStore];
  } else if (storeMap[lookupKey]) {
    result.store = storeMap[lookupKey];
  } else if (result.store) {
    result.store = result.store.replace(/(^\w|\s\w|-\w)/g, l => l.toUpperCase());
  }

  // 4. CARD MAPPING
  const cardMap = {
    "city": "Citi", "citi": "Citi", "greenstate": "GreenState", 
    "green state": "GreenState", "discover": "Discover", 
    "costco": "Costco", "verizon": "Verizon"
  };

  for (let key in cardMap) {
    if (lower.includes(key)) {
      result.card = cardMap[key];
      break; 
    }
  }

  return result;
}
