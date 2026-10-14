// ===================================================================
// RUN LOCK - One Egnyte extraction at a time (per workbook)
// ===================================================================
// Add as a new script file named "RunLock" next to Code.gs.
//
// Code.gs keeps the real logic in three renamed functions:
//   processEgnyteExtractionCore_
//   executeDeliveredSearchCore_
//   processReceivedWithOverwriteCore_
//
// Dialog.html keeps calling the original names below, which wrap the
// Core_ functions in a script lock so only one extraction writes to the
// workbook at a time. A second user either waits briefly or gets a
// "currently running" message instead of colliding.
// ===================================================================

// ===================================================================
// CONFIGURATION
// ===================================================================

const RUN_LOCK_CONFIG = {
    WAIT_MS: 20000,            // How long a second run waits for the lock before giving up
    STATUS_TTL_SECONDS: 420,   // Status auto-expires after 7 min (Apps Script max run is 6 min)
    STATUS_KEY: 'EGNYTE_RUN_STATUS'
  };
  
  // ===================================================================
  // LOCKED WRAPPERS (called from Dialog.html via google.script.run)
  // ===================================================================
  
  function processEgnyteExtraction(directLink, deliveryType, destinationTab) {
    return withRunLock_(
      `${deliveryType || 'Extraction'} on tab "${destinationTab}"`,
      () => processEgnyteExtractionCore_(directLink, deliveryType, destinationTab)
    );
  }
  
  function executeDeliveredSearch(folderId, searchListData, destinationTab, startTime) {
    const start = startTime || Date.now();
    return withRunLock_(
      `Delivered search on tab "${destinationTab}"`,
      () => executeDeliveredSearchCore_(folderId, searchListData, destinationTab, start)
    );
  }
  
  function processReceivedWithOverwrite(selectedFileNames, allFilesData, folderPath, destinationTab, startTime) {
    // Dialog.html does not pass startTime here, so time the overwrite step
    // itself (fixes "Execution Time: NaNs" on the Received report)
    const start = startTime || Date.now();
    return withRunLock_(
      `Received overwrite on tab "${destinationTab}"`,
      () => processReceivedWithOverwriteCore_(selectedFileNames, allFilesData, folderPath, destinationTab, start)
    );
  }
  
  // ===================================================================
  // LOCK WRAPPER
  // ===================================================================
  
  function withRunLock_(label, fn) {
    const lock = LockService.getScriptLock();
  
    if (!lock.tryLock(RUN_LOCK_CONFIG.WAIT_MS)) {
      const info = readRunStatus_();
      Logger.log(`Run lock busy - could not start ${label}. Current run: ${JSON.stringify(info)}`);
      // Same { success:false, message } shape the dialog already handles:
      // it shows the message and puts the form back
      return { success: false, message: buildBusyMessage_(info) };
    }
  
    try {
      writeRunStatus_(label);
      Logger.log(`Run lock acquired: ${label}`);
      return fn();
    } finally {
      try {
        CacheService.getScriptCache().remove(RUN_LOCK_CONFIG.STATUS_KEY);
      } finally {
        lock.releaseLock();
        Logger.log(`Run lock released: ${label}`);
      }
    }
  }
  
  // ===================================================================
  // STATUS (shown in the dialog banner and the busy message)
  // ===================================================================
  
  function writeRunStatus_(label) {
    const info = {
      label: label,
      user: getRunUser_(),
      startedAt: Date.now()
    };
    CacheService.getScriptCache().put(
      RUN_LOCK_CONFIG.STATUS_KEY,
      JSON.stringify(info),
      RUN_LOCK_CONFIG.STATUS_TTL_SECONDS
    );
  }
  
  function readRunStatus_() {
    const raw = CacheService.getScriptCache().get(RUN_LOCK_CONFIG.STATUS_KEY);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch (error) {
      return null;
    }
  }
  
  // Google only reveals another user's email when both accounts are in the
  // same Workspace domain; otherwise this returns "Another user"
  function getRunUser_() {
    try {
      const email = Session.getActiveUser().getEmail();
      return email ? email.split('@')[0] : 'Another user';
    } catch (error) {
      return 'Another user';
    }
  }
  
  function formatStartTime_(startedAt) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(new Date(startedAt), tz, 'h:mm a');
  }
  
  // Called by the dialog banner every few seconds.
  // The lock itself is the source of truth: if the lock is free, nothing is
  // running, even if a status entry was left behind by a run that was killed.
  function getRunStatus() {
    const lock = LockService.getScriptLock();
  
    if (lock.tryLock(0)) {
      try {
        CacheService.getScriptCache().remove(RUN_LOCK_CONFIG.STATUS_KEY); // clear any leftover status
      } finally {
        lock.releaseLock();
      }
      return { busy: false };
    }
  
    const info = readRunStatus_();
    return {
      busy: true,
      message: buildBusyMessage_(info, true)
    };
  }
  
  function buildBusyMessage_(info, forBanner) {
    const tail = forBanner
      ? 'The Extract button unlocks automatically when it finishes.'
      : 'Please try again in a minute.';
  
    if (!info) {
      return `⏳ Another extraction is in progress. ${tail}`;
    }
    return `⏳ ${info.user} is running: ${info.label} (started ${formatStartTime_(info.startedAt)}). ${tail}`;
  }