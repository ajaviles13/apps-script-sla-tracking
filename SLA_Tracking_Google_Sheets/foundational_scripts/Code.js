// ===================================================================
// CONFIGURATION - Credentials
// ===================================================================

const config_props = PropertiesService.getScriptProperties();

const CONFIG = {
  CLIENT_ID: config_props.getProperty("CLIENT_ID"), 
  CLIENT_SECRET: config_props.getProperty("CLIENT_SECRET"),
  USERNAME: config_props.getProperty("USERNAME"),
  PASSWORD: config_props.getProperty("PASSWORD"), 
  GRANT_TYPE: "password",
  SCOPE: "Egnyte.filesystem",
  EGNYTE_DOMAIN: "auratrans.egnyte.com",
  
  // Template Type: "PDF_Type" or "Word_Type"
  // PDF_Type: Standard template without isUrgent column
  // Word_Type: Template with isUrgent column in Column J
  TEMPLATE_TYPE: "Word_Type",  // Change to "Word_Type" as needed
  
  // File extensions to remove from filenames
  FILE_EXTENSIONS_TO_REMOVE: [".pdf", ".docx", ".txt"],
  
  // API Configuration
  API_CONFIG: {
    list_content: true,
    list_custom_metadata: true,
    count: 3000,  // Updated from 100 to 3000
    offset: 0,
    sort_by: "last_modified",
    sort_direction: "ascending",  // Updated from ascending to descending
    include_perm: true
  }
};

// ===================================================================
// COLUMN MAPPING HELPER - Returns column indices based on template type
// ===================================================================

function getColumnMapping() {
  const isPDFType = CONFIG.TEMPLATE_TYPE === "PDF_Type";
  
  return {
    // Common columns (same for both types)
    FILE_NAME: 2,           // Column B
    DATE_RECEIVED: 5,       // Column E
    TIME_RECEIVED: 6,       // Column F
    DATE_DELIVERED: 7,      // Column G
    TIME_DELIVERED: 8,      // Column H
    
    // Template-specific columns
    IS_URGENT: isPDFType ? null : 10,  // Column J (Word_Type only)
    QA_DELIVERY_BY: isPDFType ? 10 : 11,  // Column J (PDF) or K (Word)
    COMMENT: isPDFType ? 11 : 12,  // Column K (PDF) or L (Word)
    STATUS: isPDFType ? 12 : 13,  // Column L (PDF) or M (Word)

    // Formula columns copied onto rows added by the script
    FORMULA_FIRST: isPDFType ? 18 : 19,  // Column R (PDF) or S (Word)
    FORMULA_COUNT: 4                     // R:U (PDF) or S:V (Word)
  };
}

// ===================================================================
// ROW PLACEMENT - FILL EMPTY FILE NAME ROWS, ADD ROWS WHEN NEEDED
// ===================================================================
// A row is "empty" when its File name cell (Column B) has no text or value.
// Received imports fill empty rows from the top of the tab down. Rows that
// already have a File name are never written over. If there aren't enough
// empty rows for the whole import, the missing rows are added at the bottom.

const ROW_CONFIG = {
  DATA_START_ROW: 2,          // Row 1 is the header
  RESET_FILL_COLUMNS: 8,      // A:H fill reset to white on added rows (same as DeleteAllRowData)
  RESET_FILL_COLOR: '#ffffff'
};

function isFileNameEmpty_(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

// All rows (1-based, top to bottom) whose File name is empty
function findEmptyFileNameRows_(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < ROW_CONFIG.DATA_START_ROW) return [];

  const values = sheet.getRange(ROW_CONFIG.DATA_START_ROW, 2, maxRows - ROW_CONFIG.DATA_START_ROW + 1, 1).getValues();
  const emptyRows = [];
  values.forEach((row, i) => {
    if (isFileNameEmpty_(row[0])) {
      emptyRows.push(ROW_CONFIG.DATA_START_ROW + i);
    }
  });
  return emptyRows;
}

// Picks `count` rows to write to: empty File name rows from the top down,
// adding rows at the bottom of the tab for any shortfall.
// Returns { rows: [row numbers, ascending], rowsAdded }
function reserveEmptyRows_(sheet, count) {
  let emptyRows = findEmptyFileNameRows_(sheet);
  let rowsAdded = 0;

  if (emptyRows.length < count) {
    rowsAdded = count - emptyRows.length;
    addRowsAtBottom_(sheet, rowsAdded);
    emptyRows = findEmptyFileNameRows_(sheet);
  }

  return { rows: emptyRows.slice(0, count), rowsAdded: rowsAdded };
}

// Adds `numRows` blank rows at the bottom of the tab. Rows are inserted
// INSIDE the sheet's existing ranges (above the last row) so filters,
// dropdowns and conditional formatting grow with them; formatting, dropdowns
// and the formula columns are then copied onto the new rows.
function addRowsAtBottom_(sheet, numRows) {
  const cols = getColumnMapping();
  const maxRows = sheet.getMaxRows();
  let firstNewRow;

  if (maxRows < ROW_CONFIG.DATA_START_ROW) {
    // Only the header exists - nothing to insert inside of
    sheet.insertRowsAfter(maxRows, numRows);
    firstNewRow = maxRows + 1;
  } else {
    const lastRowHasFile = !isFileNameEmpty_(sheet.getRange(maxRows, cols.FILE_NAME).getValue());
    sheet.insertRowsBefore(maxRows, numRows);
    // The original last row is now at maxRows + numRows
    if (lastRowHasFile) {
      // That last row holds a file - move it back above the new blank rows
      const shiftedRow = maxRows + numRows;
      sheet.moveRows(sheet.getRange(`${shiftedRow}:${shiftedRow}`), maxRows);
      firstNewRow = maxRows + 1;
    } else {
      firstNewRow = maxRows;
    }
  }

  copyRowSetupToNewRows_(sheet, firstNewRow, numRows, cols);
  Logger.log(`Added ${numRows} row(s) to "${sheet.getName()}" at rows ${firstNewRow}-${firstNewRow + numRows - 1}`);
}

// Copies formatting, dropdowns and the formula columns from a neighbouring
// row onto newly added rows, then resets the A:H fill to white.
function copyRowSetupToNewRows_(sheet, firstNewRow, numRows, cols) {
  const lastNewRow = firstNewRow + numRows - 1;
  let sourceRow = firstNewRow - 1;
  if (sourceRow < ROW_CONFIG.DATA_START_ROW) {
    sourceRow = lastNewRow + 1;
  }
  if (sourceRow > sheet.getMaxRows()) {
    Logger.log('No data row to copy formatting from - new rows left unformatted');
    return;
  }

  const lastCol = sheet.getMaxColumns();
  const source = sheet.getRange(sourceRow, 1, 1, lastCol);
  const target = sheet.getRange(firstNewRow, 1, numRows, lastCol);
  source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

  if (cols.FORMULA_FIRST + cols.FORMULA_COUNT - 1 <= lastCol) {
    sheet.getRange(sourceRow, cols.FORMULA_FIRST, 1, cols.FORMULA_COUNT)
      .copyTo(
        sheet.getRange(firstNewRow, cols.FORMULA_FIRST, numRows, cols.FORMULA_COUNT),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );
  }

  sheet.getRange(firstNewRow, 1, numRows, ROW_CONFIG.RESET_FILL_COLUMNS)
    .setBackground(ROW_CONFIG.RESET_FILL_COLOR);
}

// Safety net: confirm every target row still has an empty File name right
// before writing. If not, stop the whole import and list the rows.
function assertRowsEmpty_(sheet, rows) {
  if (rows.length === 0) return;
  const lastRow = rows[rows.length - 1];
  const values = sheet.getRange(1, 2, lastRow, 1).getValues();
  const usedRows = [];
  rows.forEach(row => {
    const value = values[row - 1][0];
    if (!isFileNameEmpty_(value)) {
      usedRows.push(`row ${row} ("${value}")`);
    }
  });
  if (usedRows.length > 0) {
    throw new Error(
      `Import stopped - nothing was written. These rows already contain a File name: ${usedRows.join(', ')}. ` +
      'Please try again; if this keeps happening, contact support.'
    );
  }
}

// Writes one value per row into `column` for the given (ascending) rows.
// Consecutive rows are written together to keep the number of calls low.
function writeColumnToRows_(sheet, rows, column, values) {
  let i = 0;
  while (i < rows.length) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) {
      j++;
    }
    sheet.getRange(rows[i], column, j - i + 1, 1)
      .setValues(values.slice(i, j + 1).map(value => [value]));
    i = j + 1;
  }
}

// "5, 9-12, 40" - compact row list for the execution log
function describeRows_(rows) {
  const parts = [];
  let i = 0;
  while (i < rows.length) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
    parts.push(i === j ? `${rows[i]}` : `${rows[i]}-${rows[j]}`);
    i = j + 1;
  }
  return parts.join(', ');
}

// ===================================================================
// CHECK IF FOLDER PATH CONTAINS "URGENT"
// ===================================================================

function isFolderUrgent(folderPath) {
  if (!folderPath) return false;
  return folderPath.toLowerCase().includes("urgent");
}

// ===================================================================
// SHOW TOAST NOTIFICATION
// ===================================================================

function showToast(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '📦 Egnyte File Extraction', 5);
}

// ===================================================================
// MENU CREATION
// ===================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Egnyte File Extraction')
    .addItem('⤵️ Extract File Metadata', 'showDialog')
    .addItem('➕ Generate Language Tab', 'showGenerateNewLanguageTab')
    //.addItem('🧮 Calculate Status and SLA', 'showCalcStatusSLA')
    .addToUi();
}

// ===================================================================
// SHOW DIALOG
// ===================================================================

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📦 Egnyte File Extraction');
}

// ===================================================================
// SHOW GENERATE NEW LANGUAGE TAB DIALOG
// ===================================================================

function showGenerateNewLanguageTab() {
  const html = HtmlService.createHtmlOutputFromFile('showGenerateNewLanguageTab')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Generate Language Tab');
}

// ===================================================================
// SHOW CALCULATE STATUS AND SLA DIALOG
// ===================================================================

function showCalcStatusSLA() {
  const html = HtmlService.createHtmlOutputFromFile('CalcStatusSLA_html')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🧮 Calculate Status and SLA');
}

// ===================================================================
// GENERATE NEW LANGUAGE TAB FROM TEMPLATE
// ===================================================================

function generateNewLanguageTab(newTabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TEMPLATE_SHEET_NAME = 'TEMPLATE'; // Confirmed: Exact match in all caps
  
  try {
    // 1. Check for Duplicate Tab Name
    const allSheetNames = getSheetNames();
    if (allSheetNames.includes(newTabName)) {
      throw new Error(`The tab name "${newTabName}" already exists. Please enter a unique name.`);
    }
    
    // 2. Get the Template Sheet
    const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
    if (!templateSheet) {
      // Specific user-friendly error if TEMPLATE is missing
      throw new Error(`Error: The source tab "${TEMPLATE_SHEET_NAME}" could not be found. Please ensure a sheet named "TEMPLATE" exists.`);
    }
    
    // 3. Duplicate the Template Sheet
    // This copies ALL properties (colors, formulas, formatting, etc.)
    const newSheet = templateSheet.copyTo(ss);
    
    // 4. Rename the new sheet
    newSheet.setName(newTabName);
    
    // 5. Move the new sheet after the template (or to the end)
    const templateIndex = ss.getSheets().findIndex(sheet => sheet.getName() === TEMPLATE_SHEET_NAME);
    if (templateIndex !== -1) {
      // Move the new sheet right after the template sheet
      ss.setActiveSheet(newSheet);
      // getSheets() is 0-indexed, moveActiveSheet is 1-indexed. Index starts from 1.
      ss.moveActiveSheet(templateIndex + 2); 
    }
    
    // 6. Return success message
    showToast(`Tab "${newTabName}" successfully generated!`);
    Logger.log(`New tab "${newTabName}" created successfully from "${TEMPLATE_SHEET_NAME}".`);
    
    return {
      success: true,
      message: `✅ Success! New tab "${newTabName}" has been generated and inserted.`
    };
    
  } catch (error) {
    Logger.log('ERROR in generateNewLanguageTab: ' + error.message);
    
    // Attempt to parse Google Sheets API errors for better user feedback
    let errorMessage = error.message;
    if (errorMessage.includes("Invalid name")) {
      errorMessage = `Invalid tab name. Please ensure the name does not contain invalid characters (e.g., *, /, ?) or is not too long. Details: ${errorMessage}`;
    }
    
    return {
      success: false,
      message: `❌ Error: ${errorMessage}`
    };
  }
}

// ===================================================================
// GET CURRENT ACTIVE SHEET NAME
// ===================================================================

function getCurrentSheetName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  return activeSheet.getName();
}

// ===================================================================
// GET SHEET NAMES
// ===================================================================

function getSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  return sheets.map(sheet => sheet.getName());
}

// ===================================================================
// GET SHEET INFO (CURRENT + ALL SHEETS) - SINGLE CALL
// ===================================================================

function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheets = ss.getSheets();
  
  return {
    currentSheetName: activeSheet.getName(),
    allSheetNames: sheets.map(sheet => sheet.getName())
  };
}

// ===================================================================
// STRIP FILE EXTENSION FROM FILENAME
// ===================================================================

function stripFileExtension(fileName) {
  if (!fileName) return fileName;
  
  let cleanedName = fileName;
  
  // Check each extension in the array
  for (const extension of CONFIG.FILE_EXTENSIONS_TO_REMOVE) {
    // Simple replace - if extension exists in filename, replace with empty string
    if (cleanedName.includes(extension)) {
      cleanedName = cleanedName.replace(extension, "");
    }
  }
  
  return cleanedName;
}

// ===================================================================
// EXTRACT FOLDER ID FROM DIRECT LINK
// ===================================================================

function extractFolderIdFromLink(link) {
  // Pattern to match folder ID in the URL
  const pattern = /\/navigate\/folder\/([a-f0-9\-]+)/;
  const match = link.match(pattern);
  
  if (match && match[1]) {
    return match[1];
  }
  
  throw new Error('Could not extract folder ID from the link. Please ensure the link is in the format: https://auratrans.egnyte.com/navigate/folder/{FOLDER_ID}');
}

// ===================================================================
// ACCESS TOKEN MANAGEMENT (Updated for Refresh Token Support)
// ===================================================================

function getAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const storedToken = props.getProperty('EGNYTE_ACCESS_TOKEN');
  const storedExpiration = props.getProperty('EGNYTE_TOKEN_EXPIRATION');
  const storedRefreshToken = props.getProperty('EGNYTE_REFRESH_TOKEN');
  
  // Check if we have a valid cached token
  if (storedToken && storedExpiration) {
    const expirationTime = parseInt(storedExpiration);
    const currentTime = new Date().getTime();
    
    // If token is still valid (with 5 minute buffer), return it
    if (currentTime < expirationTime - (5 * 60 * 1000)) {
      Logger.log('Using cached access token');
      return storedToken;
    }
  }
  
  // Token expired - try refresh token flow first if available
  if (storedRefreshToken) {
    Logger.log('Access token expired - attempting refresh token flow');
    try {
      return refreshAccessToken(storedRefreshToken);
    } catch (error) {
      Logger.log(`Refresh token flow failed: ${error.message}`);
      Logger.log('Falling back to password authentication');
      // Clear invalid refresh token
      props.deleteProperty('EGNYTE_REFRESH_TOKEN');
      // Fall through to password authentication below
    }
  }
  
  // Need to get a new token via password grant
  Logger.log('Getting new access token via password grant');
  const tokenUrl = `https://${CONFIG.EGNYTE_DOMAIN}/puboauth/token`;
  
  const payload = {
    client_id: CONFIG.CLIENT_ID,
    client_secret: CONFIG.CLIENT_SECRET,
    username: CONFIG.USERNAME,
    password: CONFIG.PASSWORD,
    grant_type: CONFIG.GRANT_TYPE,
    scope: CONFIG.SCOPE
  };
  
  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const accessToken = data.access_token;
      const expiresIn = data.expires_in; // in seconds
      const refreshToken = data.refresh_token; // May or may not be present
      
      // Calculate expiration time
      const expirationTime = new Date().getTime() + (expiresIn * 1000);
      
      // Store token and expiration
      props.setProperty('EGNYTE_ACCESS_TOKEN', accessToken);
      props.setProperty('EGNYTE_TOKEN_EXPIRATION', expirationTime.toString());
      
      // Store refresh token if provided
      if (refreshToken) {
        props.setProperty('EGNYTE_REFRESH_TOKEN', refreshToken);
        Logger.log('Refresh token received and stored');
      } else {
        Logger.log('No refresh token in response (may not be supported yet)');
      }
      
      Logger.log('New access token obtained and cached');
      return accessToken;
      
    } else {
      const errorData = JSON.parse(response.getContentText());
      throw new Error(`Authentication failed (${responseCode}): ${JSON.stringify(errorData)}`);
    }
    
  } catch (error) {
    throw new Error(`Failed to get access token: ${error.message}`);
  }
}

// ===================================================================
// REFRESH ACCESS TOKEN USING REFRESH TOKEN
// ===================================================================

function refreshAccessToken(refreshToken) {
  Logger.log('Refreshing access token using refresh token');
  const tokenUrl = `https://${CONFIG.EGNYTE_DOMAIN}/puboauth/token`;
  
  const payload = {
    client_id: CONFIG.CLIENT_ID,
    client_secret: CONFIG.CLIENT_SECRET,
    refresh_token: refreshToken,
    grant_type: 'refresh_token'
  };
  
  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const accessToken = data.access_token;
      const expiresIn = data.expires_in;
      const newRefreshToken = data.refresh_token; // New refresh token
      
      // Calculate expiration time
      const expirationTime = new Date().getTime() + (expiresIn * 1000);
      
      // Store new tokens
      const props = PropertiesService.getScriptProperties();
      props.setProperty('EGNYTE_ACCESS_TOKEN', accessToken);
      props.setProperty('EGNYTE_TOKEN_EXPIRATION', expirationTime.toString());
      
      // Store new refresh token (refresh tokens can rotate)
      if (newRefreshToken) {
        props.setProperty('EGNYTE_REFRESH_TOKEN', newRefreshToken);
        Logger.log('New refresh token received and stored');
      }
      
      Logger.log('Access token refreshed successfully');
      return accessToken;
      
    } else {
      const errorData = JSON.parse(response.getContentText());
      throw new Error(`Token refresh failed (${responseCode}): ${JSON.stringify(errorData)}`);
    }
    
  } catch (error) {
    throw new Error(`Failed to refresh access token: ${error.message}`);
  }
}

// ===================================================================
// GET FOLDER CONTENTS FROM EGNYTE (WITH OFFSET FOR PAGINATION)
// ===================================================================

function getFolderContents(folderId, offset = 0) {
  const accessToken = getAccessToken();
  const apiUrl = `https://${CONFIG.EGNYTE_DOMAIN}/pubapi/v1/fs/ids/folder/${folderId}`;
  
  // Build query parameters with dynamic offset
  const params = [];
  for (const [key, value] of Object.entries(CONFIG.API_CONFIG)) {
    if (key === 'offset') {
      params.push(`${key}=${offset}`);
    } else {
      params.push(`${key}=${encodeURIComponent(value)}`);
    }
  }
  const queryString = params.join('&');
  const fullUrl = `${apiUrl}?${queryString}`;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${accessToken}`
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(fullUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      return data;
      
    } else {
      const errorData = JSON.parse(response.getContentText());
      throw new Error(`API request failed (${responseCode}): ${JSON.stringify(errorData)}`);
    }
    
  } catch (error) {
    throw new Error(`Failed to get folder contents: ${error.message}`);
  }
}

// ===================================================================
// CONVERT EPOCH MILLISECONDS TO EASTERN TIME (EST/EDT)
// ===================================================================

function convertEpochToEST(epochMilliseconds) {
  // Convert epoch milliseconds to Date object
  const date = new Date(epochMilliseconds);
  
  // Use Utilities.formatDate() with America/New_York timezone
  // This automatically handles EST (GMT-5) vs EDT (GMT-4) based on date
  const timeZone = 'America/New_York';
  
  // Format date as MM/DD/YYYY
  const dateFormatted = Utilities.formatDate(date, timeZone, 'MM/dd/yyyy');
  
  // Format time as H:MM:SS AM/PM
  // We need to build this manually since formatDate doesn't have single-digit hour format
  const hourMinSecAMPM = Utilities.formatDate(date, timeZone, 'HH:mm:ss a');
  const parts = hourMinSecAMPM.split(' ');
  const timeParts = parts[0].split(':');
  const ampm = parts[1];
  
  // Convert 24-hour to 12-hour format with single digit for hours 1-9
  let hours = parseInt(timeParts[0], 10);
  const minutes = timeParts[1];
  const seconds = timeParts[2];
  
  if (hours === 0) {
    hours = 12; // Midnight
  } else if (hours > 12) {
    hours = hours - 12; // PM hours
  }
  // For hours 1-12, keep as is (no zero padding)
  
  const timeFormatted = `${hours}:${minutes}:${seconds} ${ampm}`;
  
  return {
    date: dateFormatted,
    time: timeFormatted
  };
}

// ===================================================================
// FORMAT SHEET DATE/TIME TO MATCH EXTRACT FORMAT
// ===================================================================

function formatSheetDateTime(dateValue, timeValue) {
  // Handle empty values
  if (!dateValue && !timeValue) {
    return { date: '', time: '' };
  }
  
  let dateFormatted = '';
  let timeFormatted = '';
  
  // Format date if present - use UTC methods to avoid timezone conversion
  if (dateValue) {
    const date = new Date(dateValue);
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const day = String(date.getUTCDate()).padStart(2, '0');
    const year = date.getUTCFullYear();
    dateFormatted = `${month}/${day}/${year}`;
  }
  
  // Format time if present - use UTC methods to avoid timezone conversion
  if (timeValue) {
    const time = new Date(timeValue);
    let hours = time.getUTCHours();
    const minutes = String(time.getUTCMinutes()).padStart(2, '0');
    const seconds = String(time.getUTCSeconds()).padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // the hour '0' should be '12'
    timeFormatted = `${hours}:${minutes}:${seconds} ${ampm}`;
  }
  
  return {
    date: dateFormatted,
    time: timeFormatted
  };
}

// ===================================================================
// WRITE DATA TO SHEET (FOR RECEIVED MODE)
// ===================================================================

function writeDataToSheet(folderData, sheetName, deliveryType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  const files = folderData.files || [];
  
  if (files.length === 0) {
    throw new Error('No files found in the folder');
  }
  
  // Get column mapping
  const cols = getColumnMapping();
  
  // Check if folder is urgent (Word_Type + Received mode only)
  const isUrgentFolder = (CONFIG.TEMPLATE_TYPE === "Word_Type" && deliveryType === 'Received') 
    ? isFolderUrgent(folderData.path) 
    : false;
  
  Logger.log(`Folder path: ${folderData.path}, isUrgentFolder: ${isUrgentFolder}`);
  
  // Fill rows whose File name is empty, from the top of the tab down,
  // adding rows at the bottom only if there aren't enough empty ones
  const placement = reserveEmptyRows_(sheet, files.length);
  const targetRows = placement.rows;
  assertRowsEmpty_(sheet, targetRows);

  // Prepare one value per file for each column written
  const fileNamesColumn = [];  // Column B
  const receivedDates = [];    // Column E
  const receivedTimes = [];    // Column F
  const deliveredDates = [];   // Column G
  const deliveredTimes = [];   // Column H
  const urgentFlags = [];      // Column J (only if Word_Type)
  
  files.forEach(file => {
    const fileNameOriginal = file.name || '';
    const fileName = stripFileExtension(fileNameOriginal);
    const uploadedTimestamp = file.uploaded || 0;
    const estDateTime = convertEpochToEST(uploadedTimestamp);
    
    // Always add File Name (Column B)
    fileNamesColumn.push(fileName);
    
    // Handle RECEIVED mode data
    if (deliveryType === 'Received') {
      // Columns E, F get data; G, H are cleared so a reused row doesn't keep
      // old delivery dates (which would hide it from Delivered mode)
      receivedDates.push(estDateTime.date); // Column E
      receivedTimes.push(estDateTime.time); // Column F
      deliveredDates.push('');              // Column G
      deliveredTimes.push('');              // Column H

      // Column J (isUrgent?) is only added if template type matches
      if (CONFIG.TEMPLATE_TYPE === "Word_Type") {
        urgentFlags.push(isUrgentFolder ? 'YES' : 'NO'); // Column J
      }
      
    } else { // Handle DELIVERED mode data
      // Columns E, F get blanks, G, H get data
      receivedDates.push('');              // Column E
      receivedTimes.push('');              // Column F
      deliveredDates.push(estDateTime.date); // Column G
      deliveredTimes.push(estDateTime.time); // Column H
      
      // We don't write to Column J in Delivered mode in this function
    }
  });

  // --- WRITE OPERATIONS ---
  // Rows may not be next to each other (empty rows are filled where they are),
  // so each column is written row-group by row-group.
  
  // 1. Column B (File Name)
  writeColumnToRows_(sheet, targetRows, cols.FILE_NAME, fileNamesColumn);
  
  // 2. Columns E and F (Date/Time Received)
  writeColumnToRows_(sheet, targetRows, cols.DATE_RECEIVED, receivedDates);
  writeColumnToRows_(sheet, targetRows, cols.TIME_RECEIVED, receivedTimes);
  
  // 3. Columns G and H (Date/Time Delivered)
  writeColumnToRows_(sheet, targetRows, cols.DATE_DELIVERED, deliveredDates);
  writeColumnToRows_(sheet, targetRows, cols.TIME_DELIVERED, deliveredTimes);
  
  // 4. Conditionally Write Column J (isUrgent?) for Word_Type/Received
  if (CONFIG.TEMPLATE_TYPE === "Word_Type" && deliveryType === 'Received') {
    writeColumnToRows_(sheet, targetRows, cols.IS_URGENT, urgentFlags);
    Logger.log(`Wrote ${targetRows.length} values to Column J (isUrgent?).`);
  }
  
  // Only Columns B, E, F, G, H and (conditionally) J are written.
  // Columns A, C, D, I, K, L, etc. are untouched, preserving any existing data or formulas.

  Logger.log(`Wrote ${targetRows.length} file(s) to "${sheetName}" rows ${describeRows_(targetRows)} (${placement.rowsAdded} row(s) added)`);

  // Return file names and how many rows had to be added to the tab
  return { fileNames: fileNamesColumn, rowsAdded: placement.rowsAdded };
}

// ===================================================================
// CHECK FOR DUPLICATES IN RECEIVED MODE
// ===================================================================

function checkForDuplicatesReceived(files, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  // Get column mapping
  const cols = getColumnMapping();
  
  // Get all file names from Column B and timestamps from E & F
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return []; // No data, no duplicates
  }
  
  const columnBData = sheet.getRange('B1:B' + lastRow).getValues();
  const columnEData = sheet.getRange('E1:E' + lastRow).getDisplayValues();
  const columnFData = sheet.getRange('F1:F' + lastRow).getDisplayValues();
  
  const duplicates = [];
  
  files.forEach(file => {
    const fileNameOriginal = file.name || '';
    const fileName = stripFileExtension(fileNameOriginal);
    const uploadedTimestamp = file.uploaded || 0;
    const estDateTime = convertEpochToEST(uploadedTimestamp);
    
    // Check if file name exists and has timestamp data
    for (let i = 0; i < columnBData.length; i++) {
      const existingFileName = columnBData[i][0];
      const existingDateReceived = columnEData[i][0];
      const existingTimeReceived = columnFData[i][0];
      
      // Check if file name matches AND either E or F has data
      if (existingFileName === fileName && (existingDateReceived || existingTimeReceived)) {
        duplicates.push({
          fileName: String(fileName),
          rowIndex: i + 1,
          currentDate: existingDateReceived || '',
          currentTime: existingTimeReceived || '',
          extractDate: String(estDateTime.date),
          extractTime: String(estDateTime.time)
        });
        break; // Found duplicate, move to next file
      }
    }
  });
  
  return duplicates;
}

// ===================================================================
// BUILD DELIVERED SEARCH LIST
// ===================================================================

function buildDeliveredSearchList(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  // Get column mapping
  const cols = getColumnMapping();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return []; // No data in sheet
  }
  
  // Get all necessary columns at once
  const columnBData = sheet.getRange('B1:B' + lastRow).getValues();
  const columnGData = sheet.getRange('G1:G' + lastRow).getValues();
  const columnHData = sheet.getRange('H1:H' + lastRow).getValues();
  
  // Get QA Delivery By column based on template type
  const qaDeliveryByColumn = cols.QA_DELIVERY_BY;
  const columnLetter = String.fromCharCode(64 + qaDeliveryByColumn); // Convert column number to letter
  const qaDeliveryData = sheet.getRange(`${columnLetter}1:${columnLetter}${lastRow}`).getValues();
  
  Logger.log(`Using column ${columnLetter} (${qaDeliveryByColumn}) for QA Delivery By (Template: ${CONFIG.TEMPLATE_TYPE})`);
  
  const searchList = [];
  
  for (let i = 0; i < columnBData.length; i++) {
    const fileName = columnBData[i][0];
    const dateDelivered = columnGData[i][0];
    const timeDelivered = columnHData[i][0];
    const qaDeliveryBy = qaDeliveryData[i][0];
    
    // Include if: Column B has value AND QA Delivery By has value AND (Column G OR H is empty)
    if (fileName && qaDeliveryBy && (!dateDelivered || !timeDelivered)) {
      searchList.push({
        fileName: String(fileName),
        rowIndex: i + 1
      });
    }
  }
  
  Logger.log(`Built search list with ${searchList.length} files`);
  return searchList;
}

// ===================================================================
// GET FILES WITH FLAG NOT SET (FOR DELIVERED REPORT)
// ===================================================================

function getFilesWithFlagNotSet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  // Get column mapping
  const cols = getColumnMapping();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return []; // No data in sheet
  }
  
  // Get all necessary columns at once
  const columnBData = sheet.getRange('B1:B' + lastRow).getValues();
  const columnEData = sheet.getRange('E1:E' + lastRow).getValues();
  const columnFData = sheet.getRange('F1:F' + lastRow).getValues();
  const columnGData = sheet.getRange('G1:G' + lastRow).getValues();
  const columnHData = sheet.getRange('H1:H' + lastRow).getValues();
  
  // Get QA Delivery By column based on template type
  const qaDeliveryByColumn = cols.QA_DELIVERY_BY;
  const columnLetter = String.fromCharCode(64 + qaDeliveryByColumn); // Convert column number to letter
  const qaDeliveryData = sheet.getRange(`${columnLetter}1:${columnLetter}${lastRow}`).getValues();
  
  Logger.log(`Checking for files with flag not set in column ${columnLetter} (${qaDeliveryByColumn})`);
  
  const flagNotSetList = [];
  
  for (let i = 0; i < columnBData.length; i++) {
    const fileName = columnBData[i][0];
    const dateReceived = columnEData[i][0];
    const timeReceived = columnFData[i][0];
    const dateDelivered = columnGData[i][0];
    const timeDelivered = columnHData[i][0];
    const qaDeliveryBy = qaDeliveryData[i][0];
    
    // Include if: 
    // - Column B has value (file name exists)
    // - Columns E AND F have values (received timestamps exist)
    // - Columns G AND H are empty (no delivered timestamps)
    // - QA Delivery By column is empty
    if (fileName && dateReceived && timeReceived && !dateDelivered && !timeDelivered && !qaDeliveryBy) {
      flagNotSetList.push(String(fileName));
    }
  }
  
  Logger.log(`Found ${flagNotSetList.length} files with flag not set`);
  return flagNotSetList;
}

// ===================================================================
// SEARCH FOLDER FOR DELIVERED FILES (WITH PAGINATION)
// ===================================================================

function searchFolderForDeliveredFiles(folderId, searchList) {
  const totalToFind = searchList.length;
  let foundFiles = [];
  let notFoundFiles = [];
  let offset = 0;
  let hasMoreFiles = true;
  let folderPath = ''; // Track folder path from first API call
  
  // Create a Map for O(1) lookup
  const searchMap = new Map();
  searchList.forEach(item => {
    searchMap.set(item.fileName, item.rowIndex);
  });
  
  Logger.log(`Starting search for ${totalToFind} files`);
  
  while (hasMoreFiles && searchMap.size > 0) {
    try {
      Logger.log(`API call with offset ${offset}, remaining files to find: ${searchMap.size}`);
      
      const folderData = getFolderContents(folderId, offset);
      const files = folderData.files || [];
      const totalCount = folderData.total_count || 0;
      
      // Capture folder path from first API call
      if (offset === 0 && folderData.path) {
        folderPath = folderData.path;
      }
      
      Logger.log(`Retrieved ${files.length} files from API, total in folder: ${totalCount}`);
      
      // Check each API file against search list
      files.forEach(file => {
        const fileNameOriginal = file.name || '';
        const fileName = stripFileExtension(fileNameOriginal);
        
        if (searchMap.has(fileName)) {
          const rowIndex = searchMap.get(fileName);
          const uploadedTimestamp = file.uploaded || 0;
          const estDateTime = convertEpochToEST(uploadedTimestamp);
          
          foundFiles.push({
            fileName: fileName,
            rowIndex: rowIndex,
            date: estDateTime.date,
            time: estDateTime.time
          });
          
          // Remove from search map
          searchMap.delete(fileName);
          Logger.log(`Found: ${fileName} (${foundFiles.length}/${totalToFind})`);
        }
      });
      
      // Check if there are more files to fetch
      offset += CONFIG.API_CONFIG.count;
      hasMoreFiles = offset < totalCount;
      
      // If we found all files, stop searching
      if (searchMap.size === 0) {
        Logger.log('All files found, stopping search');
        hasMoreFiles = false;
      }
      
    } catch (error) {
      Logger.log(`Error during pagination at offset ${offset}: ${error.message}`);
      throw error;
    }
  }
  
  // Any files remaining in searchMap were not found
  searchMap.forEach((rowIndex, fileName) => {
    notFoundFiles.push({
      fileName: fileName,
      rowIndex: rowIndex
    });
  });
  
  Logger.log(`Search complete. Found: ${foundFiles.length}, Not found: ${notFoundFiles.length}`);
  
  return {
    found: foundFiles,
    notFound: notFoundFiles,
    folderPath: folderPath
  };
}

// ===================================================================
// UPDATE SHEET WITH DELIVERED FILES - WITH STATUS AND SLA CALCULATION
// ===================================================================
// Returns the file names that were SKIPPED because their row no longer
// holds the expected file name (tab sorted or rows deleted between the
// search-list build and this write). The caller reports them as Not Found.

function updateSheetWithDeliveredFiles(foundFiles, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  if (foundFiles.length === 0) {
    Logger.log('No files to update');
    return [];
  }
  
  // Get column mapping
  const cols = getColumnMapping();
  
  // Determine template type for calculateStatus and calculateSLA
  const templateType = CONFIG.TEMPLATE_TYPE === "PDF_Type" ? "PDF" : "Word";
  
  // Step 1: Read all necessary data upfront in ONE batch operation
  const rowIndices = foundFiles.map(file => file.rowIndex);
  const minRow = Math.min(...rowIndices);
  const maxRow = Math.max(...rowIndices);
  
  // Read column B (File Name) for all rows - used to confirm each row still holds the expected file
  const fileNameData = sheet.getRange(minRow, cols.FILE_NAME, maxRow - minRow + 1, 1).getValues();

  // Read columns E (Date Received) and F (Time Received) for all rows
  const dateReceivedData = sheet.getRange(minRow, cols.DATE_RECEIVED, maxRow - minRow + 1, 1).getValues();
  const timeReceivedData = sheet.getRange(minRow, cols.TIME_RECEIVED, maxRow - minRow + 1, 1).getValues();
  
  // Read Comment column (K for PDF, L for Word) - used to label NA results
  const commentData = sheet.getRange(minRow, cols.COMMENT, maxRow - minRow + 1, 1).getDisplayValues();
  
  // Read column J (isUrgent) if Word_Type
  let isUrgentData = [];
  if (CONFIG.TEMPLATE_TYPE === "Word_Type") {
    isUrgentData = sheet.getRange(minRow, cols.IS_URGENT, maxRow - minRow + 1, 1).getValues();
  }
  
  // Files whose row moved since the search list was built (skipped, reported as Not Found)
  const movedFiles = [];

  // Step 2: Process each file - calculate Status, SLA, and write data
  foundFiles.forEach(file => {
    // Get the offset from minRow to access the correct array index
    const offset = file.rowIndex - minRow;

    // Step 2.0: Row safety check - skip if this row no longer holds the expected file
    const fileNameInRow = String(fileNameData[offset][0]);
    if (fileNameInRow !== file.fileName) {
      Logger.log(`Row ${file.rowIndex} skipped: expected "${file.fileName}" but found "${fileNameInRow}" (tab sorted or rows changed)`);
      movedFiles.push(file.fileName);
      return;
    }

    // Get existing received data for this row
    const dateReceived = dateReceivedData[offset][0];
    const timeReceived = timeReceivedData[offset][0];
    
    // Get isUrgent if Word_Type
    let isUrgent = "";
    if (CONFIG.TEMPLATE_TYPE === "Word_Type") {
      isUrgent = isUrgentData[offset][0] || "";
    }
    
    // Get Comment for this row (determines Molina vs. Auracom NA label)
    const comment = commentData[offset][0] || "";
    
    // Step 2.1: Calculate Status
    let statusValue;
    try {
      statusValue = calculateStatus(
        dateReceived, 
        timeReceived, 
        file.date,  // new delivered date
        file.time,  // new delivered time
        templateType,
        isUrgent
      );
      
      // Check if calculateStatus returned an error/awaiting message
      if (statusValue && (
        statusValue.includes("Invalid date") || 
        statusValue.includes("Awaiting") || 
        statusValue.includes("Error calculating")
      )) {
        statusValue = "❓Error - Invalid date or time format";
      }
    } catch (error) {
      Logger.log(`Error calculating status for ${file.fileName}: ${error.message}`);
      statusValue = "❓Error - Invalid date or time format";
    }
    
    // Step 2.2: Calculate SLA (using the calculated Status value)
    let slaValue;
    try {
      slaValue = calculateSLA(
        templateType,
        isUrgent,  // For PDF this will be empty string, for Word it will be YES/NO
        statusValue,  // The calculated Status value
        comment  // Comment value - only used to label NA results
      );
      
      // Check if calculateSLA returned an error or empty
      if (!slaValue || (typeof slaValue === 'string' && slaValue.trim() === '')) {
        slaValue = "❓calc failed";
      }
    } catch (error) {
      Logger.log(`Error calculating SLA for ${file.fileName}: ${error.message}`);
      slaValue = "❓calc failed";
    }
    
    // Step 2.3: Write Status (column L or M)
    sheet.getRange(file.rowIndex, cols.STATUS).setValue(statusValue);
    
    // Step 2.4: Write G:I together (Date Delivered, Time Delivered, SLA Met)
    sheet.getRange(file.rowIndex, cols.DATE_DELIVERED, 1, 3).setValues([[file.date, file.time, slaValue]]);
  });
  
  Logger.log(`Updated ${foundFiles.length - movedFiles.length} rows with delivered timestamps, SLA, and status (${movedFiles.length} skipped - row moved)`);
  return movedFiles;
}

// ===================================================================
// MAIN PROCESS FUNCTION
// ===================================================================
function processEgnyteExtraction(directLink, deliveryType, destinationTab) {
  try {
    // Add at the very beginning of processEgnyteExtraction
    const startTime = new Date().getTime();

    Logger.log('processEgnyteExtraction started - ' + deliveryType + ' mode');
    
    // Validate inputs
    if (!directLink || directLink.trim() === '') {
      throw new Error('Please enter a Direct Link');
    }
    
    if (!deliveryType || (deliveryType !== 'Received' && deliveryType !== 'Delivered')) {
      throw new Error('Please select either "Received" or "Delivered"');
    }
    
    if (!destinationTab || destinationTab.trim() === '') {
      throw new Error('Please select a Destination Tab');
    }
    
    // Extract folder ID
    const folderId = extractFolderIdFromLink(directLink.trim());
    Logger.log('Folder ID: ' + folderId);
    
    // Handle Delivered mode with new logic
    if (deliveryType === 'Delivered') {
      Logger.log('Starting Delivered mode with new search logic');
      
      // Build search list
      const searchList = buildDeliveredSearchList(destinationTab);
      
      // Check for files with flag not set (for skipped count or special report)
      const flagNotSetFiles = getFilesWithFlagNotSet(destinationTab);
      
      if (searchList.length === 0) {
        // No files with QA flag set
        // Check if there are files that need the flag set
        if (flagNotSetFiles.length > 0) {
          // Get folder data for folder path
          const folderId = extractFolderIdFromLink(directLink.trim());
          const folderData = getFolderContents(folderId);

          // Calculate execution time
          const elapsedTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);
          
          return {
            success: true,
            showDeliveredFlagNotSetReport: true,
            folderPath: folderData.path || '',
            flagNotSetFiles: flagNotSetFiles,
            totalFiles: 0,
            executionTime: elapsedTime
          };
        }
        
        // No files need processing at all
        // Calculate execution time
        const elapsedTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);
        return {
          success: true,
          message: 'ℹ️ No files need delivery timestamps (all files either have timestamps or are not marked for QA delivery)',
          filesProcessed: 0,
          executionTime: elapsedTime
        };
      }
      
      // Return initial status to trigger progress dialog with counts
      return {
        success: true,
        showDeliveredProgress: true,
        totalToFind: searchList.length,
        totalSkipped: flagNotSetFiles.length,
        folderId: folderId,
        destinationTab: destinationTab,
        searchListData: JSON.stringify(searchList),
        startTime: startTime 
      };
    }
    
    // Handle Received mode (existing logic)
    if (deliveryType === 'Received') {
      // Get folder contents
      const folderData = getFolderContents(folderId);
      const files = folderData.files || [];
      Logger.log('Files retrieved: ' + files.length);
      
      if (files.length === 0) {
        throw new Error('No files found in the folder');
      }
      
      const duplicates = checkForDuplicatesReceived(files, destinationTab);
      Logger.log('Duplicates found: ' + duplicates.length);
      
      if (duplicates.length > 0) {
        // Simplify files data to only essential fields for serialization
        const simplifiedFiles = files.map(f => ({
          name: f.name || '',
          uploaded: f.uploaded || 0
        }));

        // Calculate execution time
        const elapsedTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);
        
        // Include folder path for urgency detection
        const response = {
          success: true,
          showDuplicateDialog: true,
          duplicates: duplicates,
          allFilesData: String(JSON.stringify(simplifiedFiles)),
          folderPath: String(folderData.path || ''),
          directLink: String(directLink),
          deliveryType: String(deliveryType),
          destinationTab: String(destinationTab),
          executionTime: elapsedTime 
        };
        return response;
      } else {
        Logger.log('No duplicates - processing normally');
        // No duplicates, process normally
        const writeResult = writeDataToSheet(folderData, destinationTab, deliveryType);
        const newFileNames = writeResult.fileNames;

        // Calculate execution time
        const elapsedTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);

        return {
          success: true,
          showReceivedReport: true,
          folderPath: folderData.path || '',
          newFiles: newFileNames,
          overwrittenFiles: [],
          totalFiles: newFileNames.length,
          rowsAdded: writeResult.rowsAdded,
          executionTime: elapsedTime
        };
      }
    }
    
  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    return {
      success: false,
      message: `❌ Error: ${error.message}`
    };
  }
}

// ===================================================================
// EXECUTE DELIVERED SEARCH (CALLED FROM FRONTEND WITH PROGRESS)
// ===================================================================
function executeDeliveredSearch(folderId, searchListData, destinationTab, startTime) {
  // Fall back to timing this step alone if the dialog doesn't pass a start time
  startTime = startTime || new Date().getTime();

  try {
    const searchList = JSON.parse(searchListData);

    // Perform the search
    const result = searchFolderForDeliveredFiles(folderId, searchList);

    // Update sheet with found files (returns files skipped because their row moved)
    let movedFiles = [];
    if (result.found.length > 0) {
      movedFiles = updateSheetWithDeliveredFiles(result.found, destinationTab) || [];
    }

    // Calculate execution time
    const elapsedTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);

    // Extract file names for report (moved rows are reported as Not Found)
    const updatedFiles = result.found
      .map(f => f.fileName)
      .filter(name => !movedFiles.includes(name));
    const notFoundFiles = result.notFound.map(f => f.fileName).concat(movedFiles);
    
    // Always return report data (whether all found or some not found)
    return {
      success: true,
      showDeliveredReport: true,
      folderPath: result.folderPath || '',
      updatedFiles: updatedFiles,
      notFoundFiles: notFoundFiles,
      totalFiles: updatedFiles.length,
      executionTime: elapsedTime
    };
    
  } catch (error) {
    Logger.log('ERROR in executeDeliveredSearch: ' + error.message);
    
    // Return partial success if some files were found before error
    return {
      success: false,
      partialSuccess: true,
      message: `⚠️ Partial update completed. Error: ${error.message}`
    };
  }
}

// ===================================================================
// PROCESS RECEIVED WITH OVERWRITE SELECTIONS
// ===================================================================
function processReceivedWithOverwrite(selectedFileNames, allFilesData, folderPath, destinationTab, startTime) {
  // Dialog.html doesn't pass a start time here, so time the overwrite step itself
  // (prevents "Execution Time: NaNs" on the Received report)
  startTime = startTime || new Date().getTime();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(destinationTab);
    
    if (!sheet) {
      throw new Error(`Sheet "${destinationTab}" not found`);
    }
    
    // Get column mapping
    const cols = getColumnMapping();
    
    // Check if folder is urgent (Word_Type only)
    const isUrgent = (CONFIG.TEMPLATE_TYPE === "Word_Type") 
      ? isFolderUrgent(folderPath) 
      : false;
    
    Logger.log(`processReceivedWithOverwrite - Folder path: ${folderPath}, isUrgent: ${isUrgent}`);
    
    const files = JSON.parse(allFilesData);
    let filesAdded = 0;
    let filesOverwritten = 0;
    let filesSkipped = 0;
    
    // Track file names for report
    const newFilesList = [];
    const overwrittenFilesList = [];
    
    // Get existing data for checking
    const lastRow = sheet.getLastRow();
    const columnBData = lastRow > 0 ? sheet.getRange('B1:B' + lastRow).getValues() : [];
    const columnEData = lastRow > 0 ? sheet.getRange('E1:E' + lastRow).getDisplayValues() : [];
    const columnFData = lastRow > 0 ? sheet.getRange('F1:F' + lastRow).getDisplayValues() : [];
    
    // Pass 1: sort files into overwrites (existing rows) and new files - no writes yet
    const overwrites = [];   // { row, fileName, estDateTime }
    const newFiles = [];     // { fileName, estDateTime }
    const queuedNewNames = new Set();

    files.forEach(file => {
      const fileNameOriginal = file.name || '';
      const fileName = stripFileExtension(fileNameOriginal);
      const uploadedTimestamp = file.uploaded || 0;
      const estDateTime = convertEpochToEST(uploadedTimestamp);

      // Check if this file is a duplicate
      let duplicateRow = -1;
      for (let i = 0; i < columnBData.length; i++) {
        if (columnBData[i][0] === fileName && (columnEData[i][0] || columnFData[i][0])) {
          duplicateRow = i + 1;
          break;
        }
      }

      if (duplicateRow !== -1) {
        // Duplicate: overwrite only if the user selected it
        if (selectedFileNames.includes(fileName)) {
          overwrites.push({ row: duplicateRow, fileName: fileName, estDateTime: estDateTime });
        } else {
          filesSkipped++;
        }
      } else if (queuedNewNames.has(fileName)) {
        // Same name appears twice in this batch (e.g. .pdf and .docx) - keep the first
        filesSkipped++;
      } else {
        queuedNewNames.add(fileName);
        newFiles.push({ fileName: fileName, estDateTime: estDateTime });
      }
    });

    // Pass 2: reserve rows for new files - empty File name rows from the top
    // down, adding rows at the bottom if needed - and confirm they're empty
    // BEFORE anything is written
    let placement = { rows: [], rowsAdded: 0 };
    if (newFiles.length > 0) {
      placement = reserveEmptyRows_(sheet, newFiles.length);
      assertRowsEmpty_(sheet, placement.rows);
    }

    // Pass 3: overwrite selected duplicates (E & F, plus isUrgent for Word_Type)
    overwrites.forEach(item => {
      sheet.getRange(item.row, cols.DATE_RECEIVED).setValue(item.estDateTime.date);
      sheet.getRange(item.row, cols.TIME_RECEIVED).setValue(item.estDateTime.time);
      if (CONFIG.TEMPLATE_TYPE === "Word_Type" && cols.IS_URGENT) {
        sheet.getRange(item.row, cols.IS_URGENT).setValue(isUrgent ? 'YES' : 'NO');
      }
      filesOverwritten++;
      overwrittenFilesList.push(item.fileName);
    });

    // Pass 4: write new files into the reserved rows (B, E, F, clear G & H,
    // and J for Word_Type)
    if (newFiles.length > 0) {
      const rows = placement.rows;
      writeColumnToRows_(sheet, rows, cols.FILE_NAME, newFiles.map(f => f.fileName));
      writeColumnToRows_(sheet, rows, cols.DATE_RECEIVED, newFiles.map(f => f.estDateTime.date));
      writeColumnToRows_(sheet, rows, cols.TIME_RECEIVED, newFiles.map(f => f.estDateTime.time));
      writeColumnToRows_(sheet, rows, cols.DATE_DELIVERED, newFiles.map(() => ''));
      writeColumnToRows_(sheet, rows, cols.TIME_DELIVERED, newFiles.map(() => ''));
      if (CONFIG.TEMPLATE_TYPE === "Word_Type" && cols.IS_URGENT) {
        writeColumnToRows_(sheet, rows, cols.IS_URGENT, newFiles.map(() => isUrgent ? 'YES' : 'NO'));
      }
      filesAdded = newFiles.length;
      newFiles.forEach(f => newFilesList.push(f.fileName));
      Logger.log(`Wrote ${filesAdded} new file(s) to "${destinationTab}" rows ${describeRows_(rows)} (${placement.rowsAdded} row(s) added)`);
    }

    // At the end, calculate elapsed time
    const elapsedTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);

    return {
      success: true,
      showReceivedReport: true,
      folderPath: folderPath,
      newFiles: newFilesList,
      overwrittenFiles: overwrittenFilesList,
      totalFiles: filesAdded + filesOverwritten,
      rowsAdded: placement.rowsAdded,
      executionTime: elapsedTime
    };
    
  } catch (error) {
    Logger.log(`Error in processReceivedWithOverwrite: ${error.message}`);
    return {
      success: false,
      message: `❌ Error: ${error.message}`
    };
  }
}