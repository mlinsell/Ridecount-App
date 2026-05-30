/**
 * Serves the initial HTML page as a template.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Ride Tracker Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Helper function to inject HTML/CSS/JS from separate files.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper to get the correct spreadsheet using SHEET_ID property.
 */
function getRideSheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) throw new Error("SHEET_ID script property is not set.");
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName('Ridecount');
  if (!sheet) throw new Error("Sheet 'Ridecount' not found.");
  return sheet;
}

/**
 * Returns a mapping of column headers to 0-based indices.
 */
function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[h.toString().trim()] = i;
  });
  return map;
}

/**
 * Fetches all log entries from the sheet.
 */
function getData() {
  const sheet = getRideSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) return [];
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map = getHeaderMap(sheet);
  
  return dataRange.map(row => {
    const dateStr = String(row[map['Date']] || '');
    const dateParts = dateStr.split('/');
    let parsedDate = null;
    if (dateParts.length === 3) {
      parsedDate = new Date(dateParts[2], dateParts[1] - 1, dateParts[0]);
    }
    
    return {
      date: dateStr,
      park: row[map['Park Name']] ? String(row[map['Park Name']]) : '',
      attraction: row[map['Attraction']] ? String(row[map['Attraction']]) : '',
      count: Number(row[map['Ride Count']]) || 0,
      parsedDate: parsedDate,
      uniqueId: Date.now() + Math.random()
    };
  })
  .filter(r => r.date !== '' && r.park !== '' && r.parsedDate && !isNaN(r.parsedDate.getTime()))
  .sort((a, b) => b.parsedDate - a.parsedDate); 
}

/**
 * Appends a new ride log to the sheet and checks for milestones.
 */
function addRide(rideData) {
  const sheet = getRideSheet();
  const map = getHeaderMap(sheet);
  
  const maxIndex = Math.max(...Object.values(map));
  const newRow = new Array(maxIndex + 1).fill("");
  
  if (map['Date'] !== undefined) newRow[map['Date']] = rideData.date;
  if (map['Park Name'] !== undefined) newRow[map['Park Name']] = rideData.park;
  if (map['Attraction'] !== undefined) newRow[map['Attraction']] = rideData.attraction;
  if (map['Ride Count'] !== undefined) newRow[map['Ride Count']] = rideData.count;
  
  sheet.appendRow(newRow);
  
  // Calculate milestones
  const allData = getData();
  const attractionEntries = allData.filter(r => r.attraction === rideData.attraction);
  
  const currentTotal = attractionEntries.reduce((sum, r) => sum + r.count, 0);
  const oldTotal = currentTotal - rideData.count;
  
  const thresholds = [10, 25, 50, 100, 250, 500, 1000];
  let crossedMilestone = null;
  
  for (let i = 0; i < thresholds.length; i++) {
    const threshold = thresholds[i];
    if (oldTotal < threshold && currentTotal >= threshold) {
      crossedMilestone = threshold;
      break; 
    }
  }

  return {
    success: true,
    data: allData,
    crossedMilestone: crossedMilestone
  };
}
