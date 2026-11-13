// ===================================================================
// RUN LOCK - One Egnyte extraction at a time (per workbook)
// ===================================================================
// Add as a script file named "RunLock" next to Code.gs.
//
// Code.gs keeps the real logic in three functions:
//   processEgnyteExtractionCore_
//   executeDeliveredSearchCore_
//   processReceivedWithOverwriteCore_
//
// Dialog.html calls the wrappers below. Each extraction takes a "run lease"
// (a short record of who is running what, stored in Script Properties).
// While a lease is active, any other extraction in this workbook - from
// another person or another browser tab - is rejected immediately and the
// dialog shows the in-use banner.
//
// What one run covers:
//   Received  - the import itself. The duplicate-selection screen is NOT
//               covered; clicking Overwrite takes a new lease.
//   Delivered - both steps (build the search list, then search and write)
//               as one run. The lease is handed from step 1 to step 2 with
//               a token, so nobody can start in between.
//
// LockService is used only for the few milliseconds it takes to check and
// write the lease, never for the length of a run, so the banner's status
// checks can't get in the way of a real extraction.
// ===================================================================

// ===================================================================
// CONFIGURATION
// ===================================================================

// ON/OFF switch. true = one extraction at a time with the in-use banner.
// false = no lock and no banner; extractions run immediately as before.
const runlock_enabled = true;

const RUN_LOCK_CONFIG = {
  LEASE_KEY: 'EGNYTE_RUN_LEASE',      // Script Property holding the active lease
  RUN_TTL_MS: 7 * 60 * 1000,          // A lease expires 7 min after it was taken or renewed
                                      // (Apps Script stops any run at 6 min)
  HANDOFF_TTL_MS: 60 * 1000,          // Between Delivered step 1 and step 2. If the dialog is
                                      // closed in that moment, the workbook frees up after 1 min
  LOCK_WAIT_MS: 5000                  // Max wait for the few-millisecond lease check itself
};

// ===================================================================
// LOCKED WRAPPERS (called from Dialog.html via google.script.run)
// ===================================================================

function processEgnyteExtraction(directLink, deliveryType, destinationTab) {
  if (!runlock_enabled) {
    return processEgnyteExtractionCore_(directLink, deliveryType, destinationTab);
  }

  const label = `${deliveryType || 'Extraction'} on tab "${destinationTab}"`;
  const lease = acquireRunLease_(label);
  if (!lease.ok) {
    return buildBusyResponse_(lease.holder);
  }

  let handedOff = false;
  try {
    const result = processEgnyteExtractionCore_(directLink, deliveryType, destinationTab);

    // Delivered step 1 succeeded and step 2 will follow: keep the lease and
    // hand its token to the dialog, which passes it to executeDeliveredSearch
    if (result && result.success && result.showDeliveredProgress) {
      handOffRunLease_(lease.token);
      result.runToken = lease.token;
      handedOff = true;
    }
    return result;
  } finally {
    if (!handedOff) {
      releaseRunLease_(lease.token);
    }
  }
}

function executeDeliveredSearch(folderId, searchListData, destinationTab, startTime, runToken) {
  const start = startTime || Date.now();
  if (!runlock_enabled) {
    return executeDeliveredSearchCore_(folderId, searchListData, destinationTab, start);
  }

  const lease = resumeRunLease_(runToken, `Delivered on tab "${destinationTab}"`);
  if (!lease.ok) {
    return buildBusyResponse_(lease.holder);
  }

  try {
    return executeDeliveredSearchCore_(folderId, searchListData, destinationTab, start);
  } finally {
    releaseRunLease_(lease.token);
  }
}

function processReceivedWithOverwrite(selectedFileNames, allFilesData, folderPath, destinationTab, startTime) {
  // Dialog.html does not pass startTime here, so time the overwrite step itself
  const start = startTime || Date.now();
  if (!runlock_enabled) {
    return processReceivedWithOverwriteCore_(selectedFileNames, allFilesData, folderPath, destinationTab, start);
  }

  const lease = acquireRunLease_(`Received overwrite on tab "${destinationTab}"`);
  if (!lease.ok) {
    return buildBusyResponse_(lease.holder, false);
  }

  try {
    return processReceivedWithOverwriteCore_(selectedFileNames, allFilesData, folderPath, destinationTab, start);
  } finally {
    releaseRunLease_(lease.token);
  }
}

// ===================================================================
// RUN LEASE
// ===================================================================

// Take the lease if nobody holds an active one. Never waits for another run.
function acquireRunLease_(label) {
  return withLeaseLock_(() => {
    const current = readActiveLease_();
    if (current) {
      Logger.log(`Run rejected - ${label}. Already running: ${current.label} (${current.user})`);
      return { ok: false, holder: current };
    }
    const lease = writeLease_({
      token: Utilities.getUuid(),
      label: label,
      user: getRunUser_(),
      startedAt: Date.now()
    }, RUN_LOCK_CONFIG.RUN_TTL_MS);
    Logger.log(`Run lease taken: ${label}`);
    return { ok: true, token: lease.token };
  });
}

// Delivered step 2: continue the lease handed over by step 1. If that lease
// has expired and nobody else is running, take a fresh one instead.
function resumeRunLease_(runToken, label) {
  return withLeaseLock_(() => {
    const current = readActiveLease_();
    if (current && runToken && current.token === runToken) {
      writeLease_(current, RUN_LOCK_CONFIG.RUN_TTL_MS); // renew for the long search
      Logger.log(`Run lease continued: ${current.label}`);
      return { ok: true, token: current.token };
    }
    if (current) {
      Logger.log(`Run rejected - ${label}. Already running: ${current.label} (${current.user})`);
      return { ok: false, holder: current };
    }
    const lease = writeLease_({
      token: Utilities.getUuid(),
      label: label,
      user: getRunUser_(),
      startedAt: Date.now()
    }, RUN_LOCK_CONFIG.RUN_TTL_MS);
    Logger.log(`Run lease taken (step 2, previous lease had expired): ${label}`);
    return { ok: true, token: lease.token };
  });
}

// Shorten the lease while the dialog moves from Delivered step 1 to step 2
function handOffRunLease_(token) {
  withLeaseLock_(() => {
    const current = readActiveLease_();
    if (current && current.token === token) {
      writeLease_(current, RUN_LOCK_CONFIG.HANDOFF_TTL_MS);
    }
  });
}

// Release only our own lease
function releaseRunLease_(token) {
  try {
    withLeaseLock_(() => {
      const current = readLease_();
      if (current && current.token === token) {
        PropertiesService.getScriptProperties().deleteProperty(RUN_LOCK_CONFIG.LEASE_KEY);
        Logger.log(`Run lease released: ${current.label}`);
      }
    });
  } catch (error) {
    // Never fail a finished run over this; the lease expires on its own
    Logger.log(`Could not release run lease (it will expire on its own): ${error.message}`);
  }
}

// Runs fn while holding the script lock for the few ms of the lease check
function withLeaseLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(RUN_LOCK_CONFIG.LOCK_WAIT_MS)) {
    throw new Error('Could not check whether another extraction is running. Please try again.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function writeLease_(lease, ttlMs) {
  const updated = Object.assign({}, lease, { expiresAt: Date.now() + ttlMs });
  PropertiesService.getScriptProperties().setProperty(RUN_LOCK_CONFIG.LEASE_KEY, JSON.stringify(updated));
  return updated;
}

function readLease_() {
  const raw = PropertiesService.getScriptProperties().getProperty(RUN_LOCK_CONFIG.LEASE_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

// The lease, or null if there is none or it has expired
function readActiveLease_() {
  const lease = readLease_();
  if (!lease || !lease.expiresAt || lease.expiresAt <= Date.now()) {
    return null;
  }
  return lease;
}

// ===================================================================
// STATUS (dialog banner and busy message)
// ===================================================================

// Called by the dialog banner every few seconds. Read-only: it never takes
// the script lock, so it can't delay or block a real extraction.
function getRunStatus() {
  if (!runlock_enabled) {
    return { busy: false };
  }
  const lease = readActiveLease_();
  if (!lease) {
    return { busy: false };
  }
  return { busy: true, message: buildBusyMessage_(lease, true) };
}

// forBanner = true when the dialog shows the message in the banner (form view);
// false on the duplicate screen, where it appears as a pop-up instead
function buildBusyResponse_(holder, forBanner) {
  return { success: false, busy: true, message: buildBusyMessage_(holder, forBanner !== false) };
}

function buildBusyMessage_(lease, forBanner) {
  const tail = forBanner
    ? 'The Extract button unlocks automatically when it finishes.'
    : 'Please try again in a minute.';

  if (!lease) {
    return `⏳ Another extraction is in progress. ${tail}`;
  }
  return `⏳ ${lease.user} is running: ${lease.label} (started ${formatStartTime_(lease.startedAt)}). ${tail}`;
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

// ===================================================================
// ADMIN - run manually from the Apps Script editor if the banner is ever
// stuck (for example after a run was stopped mid-way). Not needed normally:
// a stuck lease clears itself within 7 minutes.
// ===================================================================

function clearRunLock() {
  PropertiesService.getScriptProperties().deleteProperty(RUN_LOCK_CONFIG.LEASE_KEY);
  Logger.log('Run lease cleared manually');
}