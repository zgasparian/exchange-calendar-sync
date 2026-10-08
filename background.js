/*
 * background.js — account & calendar management for an on-premises
 * Exchange server.
 *
 * Runs as a normal, unprivileged WebExtension background page (manifest.json
 * loads calendar/ews.js first, into the same global scope, so EwsClient /
 * ewsItemToSimple / simpleToEwsItemXml are already available here as
 * plain globals).
 *
 * Two storage concepts, kept separate because one EWS login can sync
 * several calendars (its primary one plus any secondary folders):
 *   - accounts[accountId]      one EWS login: {url, username, displayName,
 *                               needsReauth}. accountId = username.toLowerCase().
 *   - calendars[calendarKey]   one synced calendar (folder): {accountId,
 *                               folderRef, folderName, syncState,
 *                               lastSyncedAt, needsReauth}.
 *                               calendarKey = `${accountId}::${folderRef.id
 *                               || folderRef.distinguishedId}`.
 *
 * Passwords live only in `passwordCache` (in memory, this session only) —
 * persisted copies go through the privileged `browser.exchangeCalendar.*`
 * credential functions (nsILoginManager), never browser.storage.local.
 *
 * This file never touches XPCOM/Thunderbird internals directly — it only
 * calls `browser.exchangeCalendar.*` to register calendars and push
 * synced events into them, and listens for `onLocalChange` to push the
 * user's own edits (including meeting-response clicks) back out to the
 * server.
 */

const SYNC_ALARM = "exchangeCalendarSync";

// accountId -> password, cached in memory only (never persisted here) so
// the webRequest.onAuthRequired listener can answer NTLM/Negotiate/Basic
// challenges without round-tripping through the Experiment API each time.
const passwordCache = new Map();

function calendarKeyFor(accountId, folderRef) {
  return `${accountId}::${folderRef.id || folderRef.distinguishedId}`;
}

async function getAccounts() {
  const { accounts } = await browser.storage.local.get("accounts");
  return accounts || {};
}

async function saveAccounts(accounts) {
  await browser.storage.local.set({ accounts });
}

async function getCalendars() {
  const { calendars } = await browser.storage.local.get("calendars");
  return calendars || {};
}

async function saveCalendars(calendars) {
  await browser.storage.local.set({ calendars });
}

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

  // Also proves the URL/credentials work, same as the old throwaway test
  // call used to, but this result is actually useful (the folder list).
  const client = new EwsClient(ewsUrl, username, password);
  const folders = await client.listCalendarFolders();
  // Needed so Thunderbird's own invitation UI (Accept/Tentative/Decline)
  // can recognize "this event invites me" — see registerCalendar() below
  // and calendar/provider.js's organizerId usage. Best-effort: null just
  // means that UI won't activate, not a connection failure.
  const ownerEmail = await client.resolveOwnEmail(username);

  const accounts = await getAccounts();
  accounts[accountId] = { url: ewsUrl, username, displayName: username, ownerEmail, needsReauth: false };
  await saveAccounts(accounts);
  await browser.exchangeCalendar.saveCredentials(accountId, username, password);

  // Primary calendar auto-syncs immediately, matching prior versions'
  // behavior; secondary calendars (if any) are opt-in via addCalendar().
  const primary = folders.find(f => f.folderRef.distinguishedId) || folders[0];
  await addCalendarInternal(accountId, primary.folderRef, primary.name);
  await ensureSyncAlarm();

  return {
    account: { accountId, username, url: ewsUrl },
    folders: await annotateFolders(accountId, folders),
  };
}

/** Cross-references EWS's folder list against what's already synced, for the settings page. */
async function annotateFolders(accountId, folders) {
  const calendars = await getCalendars();
  const syncedKeys = new Set(Object.keys(calendars).filter(k => calendars[k].accountId === accountId));
  return folders.map(f => ({
    folderRef: f.folderRef,
    name: f.name,
    calendarKey: calendarKeyFor(accountId, f.folderRef),
    synced: syncedKeys.has(calendarKeyFor(accountId, f.folderRef)),
  }));
}

async function listFolders(accountId) {
  const accounts = await getAccounts();
  const account = accounts[accountId];
  if (!account) {
    throw new Error(`Unknown account ${accountId}`);
  }
  const client = new EwsClient(account.url, account.username, passwordCache.get(accountId));
  const folders = await client.listCalendarFolders();
  return annotateFolders(accountId, folders);
}

async function addCalendarInternal(accountId, folderRef, folderName) {
  const calendarKey = calendarKeyFor(accountId, folderRef);
  const calendars = await getCalendars();
  calendars[calendarKey] = {
    accountId,
    folderRef,
    folderName,
    syncState: null,
    lastSyncedAt: null,
    needsReauth: false,
  };
  await saveCalendars(calendars);

  const accounts = await getAccounts();
  await browser.exchangeCalendar.registerCalendar(
    calendarKey,
    `${accounts[accountId].displayName} — ${folderName}`,
    accounts[accountId].ownerEmail
  );
  await syncCalendar(calendarKey); // isInitialSync inside suppresses notifications for this first pass
  return calendarKey;
}

async function removeCalendar(calendarKey) {
  const calendars = await getCalendars();
  if (!calendars[calendarKey]) {
    return;
  }
  await browser.exchangeCalendar.unregisterCalendar(calendarKey);
  delete calendars[calendarKey];
  await saveCalendars(calendars);
}

async function removeAccount(accountId) {
  const accounts = await getAccounts();
  if (!accounts[accountId]) {
    return;
  }
  const calendars = await getCalendars();
  for (const calendarKey of Object.keys(calendars)) {
    if (calendars[calendarKey].accountId === accountId) {
      await removeCalendar(calendarKey);
    }
  }
  await browser.exchangeCalendar.deleteCredentials(accountId);
  passwordCache.delete(accountId);
  delete accounts[accountId];
  await saveAccounts(accounts);
}

async function syncCalendar(calendarKey) {
  const calendars = await getCalendars();
  const calendar = calendars[calendarKey];
  if (!calendar) {
    return;
  }
  const accounts = await getAccounts();
  const account = accounts[calendar.accountId];
  const password = passwordCache.get(calendar.accountId);
  if (!account || !password) {
    return;
  }
  const client = new EwsClient(account.url, account.username, password);
  // Suppresses notifications on a calendar's very first sync — otherwise
  // connecting an account with months of history would fire a desktop
  // notification for every single existing meeting.
  const isInitialSync = !calendar.syncState;
  try {
    let syncState = calendar.syncState;
    let moreAvailable = true;
    while (moreAvailable) {
      const result = await client.syncFolderItems(calendar.folderRef, syncState);
      if (result.changes.length) {
        await browser.exchangeCalendar.applyRemoteChanges(calendarKey, result.changes);
        if (!isInitialSync) {
          notifyForChanges(calendar, result.changes);
        }
      }
      syncState = result.syncState;
      moreAvailable = result.moreAvailable;
    }
    calendar.syncState = syncState;
    calendar.needsReauth = false;
    calendar.lastSyncedAt = Date.now();
    await saveCalendars(calendars);
  } catch (e) {
    if (isAuthError(e)) {
      calendar.needsReauth = true;
      await saveCalendars(calendars);
      notifyReauthRequired(account);
      return;
    }
    console.error(`exchangeCalendar: sync failed for ${calendarKey}`, e);
  }
}

function notifyForChanges(calendar, changes) {
  for (const change of changes) {
    if (change.op === "delete") {
      continue;
    }
    if (change.item.isCancelled) {
      notify(`Meeting cancelled — ${calendar.folderName}`, change.item.title || "(no title)");
    } else if (change.op === "create") {
      notify(`New meeting invite — ${calendar.folderName}`, change.item.title || "(no title)");
    }
  }
}

function notify(title, message) {
  browser.notifications
    .create({ type: "basic", iconUrl: "icons/icon-48.png", title, message })
    .catch(() => {});
}

function isAuthError(e) {
  return /\b401\b/.test(e.message || "") || (e.name === "EwsSoapFault" && e.responseCode === "ErrorAccessDenied");
}

async function syncAll() {
  const calendars = await getCalendars();
  for (const calendarKey of Object.keys(calendars)) {
    if (!calendars[calendarKey].needsReauth) {
      await syncCalendar(calendarKey);
    }
  }
}

function notifyReauthRequired(account) {
  notify("Exchange Calendar Sync", `Sign-in failed for ${account.username}. Open the add-on's settings to re-enter your password.`);
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

/**
 * Upgrading from a pre-0.4.0 version: `accounts[accountId]` used to carry
 * its own syncState/lastSyncedAt directly (one account == one, always-
 * primary, calendar) — there was no separate `calendars` store at all.
 * Synthesizes the missing `calendars[calendarKey]` entry for any account
 * that predates that split, reusing its old sync token so this resumes
 * with a normal incremental sync rather than a full resync. Pairs with
 * the matching adoption fix in provider.js's `id` setter, which re-keys
 * that account's existing (already-populated) Thunderbird calendar under
 * the same calendarKey instead of us ending up with a duplicate.
 */
async function migrateLegacyAccounts() {
  const accounts = await getAccounts();
  const calendars = await getCalendars();
  let changed = false;
  for (const [accountId, account] of Object.entries(accounts)) {
    if (Object.values(calendars).some(c => c.accountId === accountId)) {
      continue;
    }
    const calendarKey = calendarKeyFor(accountId, PRIMARY_CALENDAR);
    calendars[calendarKey] = {
      accountId,
      folderRef: PRIMARY_CALENDAR,
      folderName: "Calendar",
      syncState: account.syncState || null,
      lastSyncedAt: account.lastSyncedAt || null,
      needsReauth: false,
    };
    changed = true;
  }
  if (changed) {
    await saveCalendars(calendars);
  }
}

/** On startup: re-fetch each account's password from the encrypted store into memory, re-arm its auth handler, and make sure every known calendar is registered. */
async function restoreOnStartup() {
  await migrateLegacyAccounts();
  const accounts = await getAccounts();
  for (const [accountId, account] of Object.entries(accounts)) {
    const credentials = await browser.exchangeCalendar.getCredentials(accountId);
    if (!credentials) {
      account.needsReauth = true;
    } else {
      passwordCache.set(accountId, credentials.password);
      registerAuthHandler(accountId, account.url, account.username);
      // Self-heals accounts connected before ownerEmail existed, so
      // existing users get the Accept/Tentative/Decline UI without
      // having to disconnect and reconnect.
      if (!account.ownerEmail) {
        const client = new EwsClient(account.url, account.username, credentials.password);
        account.ownerEmail = await client.resolveOwnEmail(account.username);
      }
    }
  }
  await saveAccounts(accounts);

  const calendars = await getCalendars();
  for (const [calendarKey, calendar] of Object.entries(calendars)) {
    const account = accounts[calendar.accountId];
    const { wasStoreReset } = await browser.exchangeCalendar.registerCalendar(
      calendarKey,
      `${account?.displayName ?? calendar.accountId} — ${calendar.folderName}`,
      account?.ownerEmail
    );
    if (wasStoreReset) {
      // The on-disk calendar store was just (re)created — e.g. upgrading
      // from a version that used the non-persistent "memory" backing, or
      // from before calendars were tracked separately from accounts — so
      // our saved delta-sync token no longer corresponds to anything on
      // disk. A normal incremental sync from it would correctly see "no
      // server-side changes" and leave the calendar empty forever; force
      // a full resync instead.
      calendar.syncState = null;
    }
  }
  await saveCalendars(calendars);
}

/** Finds which attendee's own RSVP changed between two SimpleEvent snapshots of the same item, if any. */
function detectRsvpChange(item, oldItem) {
  if (!oldItem) {
    return null;
  }
  for (const newAttendee of item.attendees || []) {
    const oldAttendee = (oldItem.attendees || []).find(a => a.email === newAttendee.email);
    if (oldAttendee && oldAttendee.status !== newAttendee.status) {
      return newAttendee.status;
    }
  }
  return null;
}

browser.alarms.onAlarm.addListener(alarm => {
  if (alarm.name === SYNC_ALARM) {
    syncAll();
  }
});

// The user edited/added/deleted an event (or clicked Accept/Decline/
// Tentative on an invite) directly in Thunderbird; push it to the server.
browser.exchangeCalendar.onLocalChange.addListener(async (calendarKey, op, item, oldItem, requestId) => {
  try {
    const calendars = await getCalendars();
    const calendar = calendars[calendarKey];
    if (!calendar) {
      throw new Error(`Unknown calendar ${calendarKey}`);
    }
    const accounts = await getAccounts();
    const account = accounts[calendar.accountId];
    const password = passwordCache.get(calendar.accountId);
    if (!account || !password) {
      throw new Error(`Unknown or signed-out account ${calendar.accountId}`);
    }
    const client = new EwsClient(account.url, account.username, password);

    let result = null;
    if (op === "delete") {
      await client.deleteEvent(item);
    } else if (op === "add") {
      result = await client.createEvent(calendar.folderRef, item);
    } else {
      const rsvpStatus = detectRsvpChange(item, oldItem);
      if (rsvpStatus && ["ACCEPTED", "DECLINED", "TENTATIVE"].includes(rsvpStatus)) {
        result = await client.respondToInvite(item, rsvpStatus);
      } else {
        result = await client.updateEvent(item);
      }
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
      const calendars = await getCalendars();
      return Object.entries(accounts).map(([accountId, a]) => ({
        accountId,
        username: a.username,
        url: a.url,
        needsReauth: a.needsReauth,
        calendars: Object.entries(calendars)
          .filter(([, c]) => c.accountId === accountId)
          .map(([calendarKey, c]) => ({
            calendarKey,
            folderName: c.folderName,
            lastSyncedAt: c.lastSyncedAt,
            needsReauth: c.needsReauth,
          })),
      }));
    }
    case "connectAccount": {
      try {
        const result = await connectAccount(message.url, message.username, message.password);
        return { ok: true, ...result };
      } catch (e) {
        console.error("exchangeCalendar: connect failed", e);
        return { ok: false, error: e.message || String(e) };
      }
    }
    case "removeAccount":
      await removeAccount(message.accountId);
      return { ok: true };
    case "listFolders": {
      try {
        return { ok: true, folders: await listFolders(message.accountId) };
      } catch (e) {
        return { ok: false, error: e.message || String(e) };
      }
    }
    case "addCalendar": {
      try {
        const calendarKey = await addCalendarInternal(message.accountId, message.folderRef, message.folderName);
        return { ok: true, calendarKey };
      } catch (e) {
        return { ok: false, error: e.message || String(e) };
      }
    }
    case "removeCalendar":
      await removeCalendar(message.calendarKey);
      return { ok: true };
    case "syncNow":
      await syncCalendar(message.calendarKey);
      return { ok: true };
    case "getSyncIntervalMinutes":
      return { minutes: await getSyncIntervalMinutes() };
    case "setSyncIntervalMinutes":
      return { minutes: await setSyncIntervalMinutes(message.minutes) };
    default:
      return undefined;
  }
});

restoreOnStartup().then(ensureSyncAlarm).then(syncAll);
