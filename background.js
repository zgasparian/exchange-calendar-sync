/*
 * background.js — account management for an on-premises Exchange server.
 *
 * Runs as a normal, unprivileged WebExtension background page (manifest.json
 * loads calendar/ews.js first, into the same global scope, so EwsClient /
 * ewsItemToSimple / simpleToEwsItemXml are already available here as plain
 * globals).
 *
 * Non-secret account metadata (EWS URL, username, display name, sync
 * state) lives in browser.storage.local. The password does not — it's
 * stored encrypted via the privileged `browser.exchangeCalendar.*`
 * credential functions (see calendar/schema.json / calendar/provider.js,
 * backed by nsILoginManager) and only ever held in memory here, per
 * account, for as long as the background page is alive.
 *
 * This file never touches XPCOM/Thunderbird internals directly — it only
 * calls `browser.exchangeCalendar.*` to register a calendar and push
 * synced events into it, and listens for `onLocalChange` to push the
 * user's own edits back out to the server.
 */

const SYNC_ALARM = "exchangeCalendarSync";

// accountId -> password, cached in memory only (never persisted here) so
// the webRequest.onAuthRequired listener can answer NTLM/Negotiate/Basic
// challenges without round-tripping through the Experiment API each time.
const passwordCache = new Map();

async function getAccounts() {
  const { accounts } = await browser.storage.local.get("accounts");
  return accounts || {};
}

async function saveAccounts(accounts) {
  await browser.storage.local.set({ accounts });
}

/**
 * EWS is a sibling virtual directory to OWA on the Exchange CAS server
 * (https://mail.company.com/owa and https://mail.company.com/EWS/Exchange.asmx
 * share a host, EWS is never *under* /owa), so whatever path the user
 * pastes — the OWA link IT actually hands out, a bare hostname, or the
 * EWS URL itself — we only keep the scheme+host+port and rebuild the EWS
 * path ourselves rather than appending onto whatever path was given.
 */
function normalizeEwsUrl(input) {
  const trimmed = input.trim();
  const withScheme = /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
  const origin = new URL(withScheme).origin;
  return `${origin}/EWS/Exchange.asmx`;
}

function registerAuthHandler(accountId, url, username) {
  const origin = new URL(url).origin + "/*";
  browser.webRequest.onAuthRequired.addListener(
    details => {
      const password = passwordCache.get(accountId);
      if (!password || details.isProxy) {
        return {};
      }
      return { authCredentials: { username, password } };
    },
    { urls: [origin] },
    ["blocking"]
  );
}

async function connectAccount(url, username, password) {
  const ewsUrl = normalizeEwsUrl(url);
  const accountId = username.toLowerCase();

  passwordCache.set(accountId, password);
  registerAuthHandler(accountId, ewsUrl, username);

  // This first sync page both proves the URL/credentials work and is the
  // start of the real initial sync — no separate throwaway "test" call.
  const client = new EwsClient(ewsUrl, username, password);
  const firstPage = await client.syncFolderItems(null);

  const accounts = await getAccounts();
  accounts[accountId] = {
    url: ewsUrl,
    username,
    displayName: username,
    syncState: firstPage.syncState,
    needsReauth: false,
    lastSyncedAt: null,
  };
  await saveAccounts(accounts);
  await browser.exchangeCalendar.saveCredentials(accountId, username, password);
  await browser.exchangeCalendar.registerCalendar(accountId, accounts[accountId].displayName);
  await ensureSyncAlarm();

  if (firstPage.changes.length) {
    await browser.exchangeCalendar.applyRemoteChanges(accountId, firstPage.changes);
  }
  accounts[accountId].lastSyncedAt = Date.now();
  await saveAccounts(accounts);
  if (firstPage.moreAvailable) {
    await syncAccount(accountId); // pages through the rest, and bumps lastSyncedAt again when done
  }
  return accounts[accountId];
}

async function removeAccount(accountId) {
  const accounts = await getAccounts();
  if (!accounts[accountId]) {
    return;
  }
  await browser.exchangeCalendar.unregisterCalendar(accountId);
  passwordCache.delete(accountId);
  delete accounts[accountId];
  await saveAccounts(accounts);
}

async function syncAccount(accountId) {
  const accounts = await getAccounts();
  const account = accounts[accountId];
  const password = passwordCache.get(accountId);
  if (!account || !password) {
    return;
  }
  const client = new EwsClient(account.url, account.username, password);
  try {
    let syncState = account.syncState;
    let moreAvailable = true;
    while (moreAvailable) {
      const result = await client.syncFolderItems(syncState);
      if (result.changes.length) {
        await browser.exchangeCalendar.applyRemoteChanges(accountId, result.changes);
      }
      syncState = result.syncState;
      moreAvailable = result.moreAvailable;
    }
    account.syncState = syncState;
    account.needsReauth = false;
    account.lastSyncedAt = Date.now();
    await saveAccounts(accounts);
  } catch (e) {
    if (isAuthError(e)) {
      account.needsReauth = true;
      await saveAccounts(accounts);
      notifyReauthRequired(account);
      return;
    }
    console.error(`exchangeCalendar: sync failed for ${accountId}`, e);
  }
}

function isAuthError(e) {
  return /\b401\b/.test(e.message || "") || e.name === "EwsSoapFault" && e.responseCode === "ErrorAccessDenied";
}

async function syncAll() {
  const accounts = await getAccounts();
  for (const accountId of Object.keys(accounts)) {
    if (!accounts[accountId].needsReauth) {
      await syncAccount(accountId);
    }
  }
}

function notifyReauthRequired(account) {
  browser.notifications
    .create({
      type: "basic",
      iconUrl: "icons/icon-48.png",
      title: "Exchange Calendar Sync",
      message: `Sign-in failed for ${account.username}. Open the add-on's settings to re-enter your password.`,
    })
    .catch(() => {});
}

const DEFAULT_SYNC_INTERVAL_MINUTES = 5;
const MIN_SYNC_INTERVAL_MINUTES = 1;

async function getSyncIntervalMinutes() {
  const { syncIntervalMinutes } = await browser.storage.local.get("syncIntervalMinutes");
  return syncIntervalMinutes || DEFAULT_SYNC_INTERVAL_MINUTES;
}

/** Re-arms the alarm to match the configured interval. Safe to call anytime — e.g. right after the user changes it — since it always clears and recreates rather than assuming the existing alarm (if any) already matches. */
async function ensureSyncAlarm() {
  const periodInMinutes = await getSyncIntervalMinutes();
  await browser.alarms.clear(SYNC_ALARM);
  browser.alarms.create(SYNC_ALARM, { periodInMinutes });
}

async function setSyncIntervalMinutes(minutes) {
  const clamped = Math.max(MIN_SYNC_INTERVAL_MINUTES, Math.round(minutes) || DEFAULT_SYNC_INTERVAL_MINUTES);
  await browser.storage.local.set({ syncIntervalMinutes: clamped });
  await ensureSyncAlarm();
  return clamped;
}

/** On startup: re-fetch each account's password from the encrypted store into memory, re-arm its auth handler, and make sure its calendar is registered. */
async function restoreAccountsOnStartup() {
  const accounts = await getAccounts();
  for (const [accountId, account] of Object.entries(accounts)) {
    const credentials = await browser.exchangeCalendar.getCredentials(accountId);
    if (!credentials) {
      account.needsReauth = true;
    } else {
      passwordCache.set(accountId, credentials.password);
      registerAuthHandler(accountId, account.url, account.username);
    }
    const { wasStoreReset } = await browser.exchangeCalendar.registerCalendar(accountId, account.displayName);
    if (wasStoreReset) {
      // The on-disk calendar store was just (re)created — e.g. upgrading
      // from a version that used the non-persistent "memory" backing — so
      // our saved delta-sync token no longer corresponds to anything on
      // disk. A normal incremental sync from it would correctly see "no
      // server-side changes" and leave the calendar empty forever; force
      // a full resync instead.
      account.syncState = null;
    }
  }
  await saveAccounts(accounts);
}

browser.alarms.onAlarm.addListener(alarm => {
  if (alarm.name === SYNC_ALARM) {
    syncAll();
  }
});

// The user edited/added/deleted an event directly in Thunderbird; push it to the EWS server.
browser.exchangeCalendar.onLocalChange.addListener(async (accountId, op, item, oldItem, requestId) => {
  try {
    const accounts = await getAccounts();
    const account = accounts[accountId];
    const password = passwordCache.get(accountId);
    if (!account || !password) {
      throw new Error(`Unknown or signed-out account ${accountId}`);
    }
    const client = new EwsClient(account.url, account.username, password);

    let result = null;
    if (op === "delete") {
      await client.deleteEvent(item);
    } else if (op === "add") {
      result = await client.createEvent(item);
    } else {
      result = await client.updateEvent(item);
    }
    await browser.exchangeCalendar.resolveLocalChange(requestId, result);
  } catch (e) {
    console.error("exchangeCalendar: failed to push local change", e);
    await browser.exchangeCalendar.rejectLocalChange(requestId, e.message || String(e));
  }
});

// Bridge for ui/options.js (a separate page, so it talks to us by message
// rather than sharing this scope directly).
browser.runtime.onMessage.addListener(async message => {
  switch (message.type) {
    case "listAccounts": {
      const accounts = await getAccounts();
      return Object.entries(accounts).map(([accountId, a]) => ({
        accountId,
        displayName: a.displayName,
        username: a.username,
        url: a.url,
        needsReauth: a.needsReauth,
        lastSyncedAt: a.lastSyncedAt || null,
      }));
    }
    case "connectAccount": {
      try {
        const account = await connectAccount(message.url, message.username, message.password);
        return { ok: true, account };
      } catch (e) {
        console.error("exchangeCalendar: connect failed", e);
        return { ok: false, error: e.message || String(e) };
      }
    }
    case "removeAccount":
      await removeAccount(message.accountId);
      return { ok: true };
    case "syncNow":
      await syncAccount(message.accountId);
      return { ok: true };
    case "getSyncIntervalMinutes":
      return { minutes: await getSyncIntervalMinutes() };
    case "setSyncIntervalMinutes":
      return { minutes: await setSyncIntervalMinutes(message.minutes) };
    default:
      return undefined;
  }
});

restoreAccountsOnStartup().then(ensureSyncAlarm).then(syncAll);
