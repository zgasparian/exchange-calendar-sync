/*
 * provider.js — the privileged "parent" half of the exchangeCalendar
 * Experiment API (see manifest.json's experiment_apis.exchangeCalendar and
 * calendar/schema.json). This is the ONLY file in this extension that runs
 * with full Thunderbird/Gecko privileges (XPCOM, ChromeUtils, Services).
 *
 * It does no networking at all — that's background.js + calendar/ews.js,
 * running as a normal unprivileged WebExtension background page (ews.js
 * talks to the on-premises Exchange server's EWS SOAP endpoint). This file
 * only ever talks to Thunderbird's calendar manager, plus nsILoginManager
 * for credential storage:
 *
 *   - registers a calICalendar type ("exchangeEwsSync") and, per *synced
 *     calendar* (one EWS login/account may sync several folders — see
 *     "calendarKey" below), one calICalendar instance backed by an
 *     internal "storage" (SQLite) calendar for persistence;
 *   - applyRemoteChanges() lets background.js push what it fetched from
 *     EWS into that store, which fires the normal onAddItem/onModifyItem/
 *     onDeleteItem notifications so open calendar views update live;
 *   - when the *user* adds/edits/deletes an event in Thunderbird's own UI,
 *     this fires onLocalChange out to background.js, which pushes it to
 *     the server and calls back resolveLocalChange()/rejectLocalChange();
 *   - saveCredentials()/getCredentials()/deleteCredentials() store an EWS
 *     login's username+password in nsILoginManager (the same encrypted
 *     store IMAP/SMTP passwords use), since background.js is unprivileged
 *     and can't reach it directly.
 *
 * Naming note: an "accountId" is one EWS login (url+username+password).
 * Since one login can sync multiple calendars (primary + secondary
 * folders), each synced calendar has its own "calendarKey" (background.js
 * mints these as `${accountId}::${folderId}`) — registerCalendar/
 * unregisterCalendar/applyRemoteChanges/onLocalChange all key off
 * calendarKey, while saveCredentials/getCredentials/deleteCredentials key
 * off accountId, since credentials belong to the login, not to any one
 * synced calendar.
 *
 * The calICalendar contract used here (promise-based addItem/modifyItem/
 * deleteItem/getItem, cal.provider.BaseClass, cal.manager.register*) was
 * verified against this machine's installed Thunderbird 156 by extracting
 * omni.ja and reading CalCalendarManager.sys.mjs, calProviderUtils.sys.mjs,
 * CalDavCalendar.sys.mjs and the (unfinished, read-only) native
 * GraphCalendar.sys.mjs — not guessed from memory. Likewise CalAlarm.sys.mjs
 * and CalDuration.sys.mjs for the reminder mapping below. Still, if a
 * future Thunderbird version changes this contract, this file is where to
 * look.
 */

(function (exports) {
  "use strict";

  const { cal } = ChromeUtils.importESModule("resource:///modules/calendar/calUtils.sys.mjs");
  const { CalEvent } = ChromeUtils.importESModule("resource:///modules/CalEvent.sys.mjs");
  const { CalAttendee } = ChromeUtils.importESModule("resource:///modules/CalAttendee.sys.mjs");
  const { CalAlarm } = ChromeUtils.importESModule("resource:///modules/CalAlarm.sys.mjs");
  const { CalDuration } = ChromeUtils.importESModule("resource:///modules/CalDuration.sys.mjs");
  const { CalRecurrenceInfo } = ChromeUtils.importESModule("resource:///modules/CalRecurrenceInfo.sys.mjs");
  const { CalRecurrenceRule } = ChromeUtils.importESModule("resource:///modules/CalRecurrenceRule.sys.mjs");

  const CALENDAR_TYPE = "exchangeEwsSync";
  // Synthetic category tag applied to cancelled meetings so the user can
  // assign it a color (e.g. red) in Thunderbird's Categories preferences —
  // Thunderbird has no native per-event color outside of categories.
  // event.status = "CANCELLED" (set below) already gets a native
  // strikethrough in calendar views with no configuration needed.
  const CANCELLED_CATEGORY = "Cancelled";

  // Synthetic (never dereferenced) origin used as the nsILoginManager lookup
  // key for an account's EWS credentials — one entry per accountId. Several
  // synced calendars (calendarKeys) can share one accountId's credentials.
  function credentialScope(accountId) {
    return { origin: `exchangecalendarsync://${encodeURIComponent(accountId)}`, httpRealm: "ews" };
  }

  async function clearStoredCredentials(accountId) {
    const { origin, httpRealm } = credentialScope(accountId);
    for (const existing of await Services.logins.searchLoginsAsync({ origin, httpRealm })) {
      await Services.logins.removeLoginAsync(existing);
    }
  }

  // calendarKey -> ExchangeEwsCalendar instance (repopulated as calendars
  // are (re)created, see the `id` setter override below).
  const calendarsByKey = new Map();

  // requestId -> {resolve, reject}, for addItem/modifyItem/deleteItem calls
  // that are waiting on background.js to finish talking to the EWS server.
  const pendingLocalChanges = new Map();

  // Set by getAPI()'s onLocalChange EventManager while background.js has a
  // listener registered (effectively always, once the add-on has started).
  let fireLocalChange = null;

  /** calIAlarm for "N minutes before start", the only shape EWS reminders use. */
  function buildReminderAlarm(minutesBeforeStart) {
    const alarm = new CalAlarm();
    alarm.related = Ci.calIAlarm.ALARM_RELATED_START;
    const duration = new CalDuration();
    duration.inSeconds = -Math.abs(minutesBeforeStart) * 60;
    alarm.offset = duration;
    alarm.action = "DISPLAY";
    return alarm;
  }

  /** Reverse of buildReminderAlarm(): reads back the "before start" alarm, if any, in minutes. */
  function getReminderMinutesBeforeStart(event) {
    const alarm = event.getAlarms().find(a => a.related === Ci.calIAlarm.ALARM_RELATED_START && a.offset);
    return alarm ? Math.round(Math.abs(alarm.offset.inSeconds) / 60) : null;
  }

  function buildDescriptionText(simple) {
    let description = simple.description || "";
    if (simple.onlineMeetingUrl && !description.includes(simple.onlineMeetingUrl)) {
      const line = `Join online meeting: ${simple.onlineMeetingUrl}`;
      description = description ? `${description}\n\n${line}` : line;
    }
    return description;
  }

  function ewsResponseToPartStat(responseType) {
    switch (responseType) {
      case "Accept":
      case "ACCEPTED":
        return "ACCEPTED";
      case "Decline":
      case "DECLINED":
        return "DECLINED";
      case "Tentative":
      case "TENTATIVE":
        return "TENTATIVE";
      default:
        return "NEEDS-ACTION";
    }
  }

  function simpleToCalEvent(simple) {
    const event = new CalEvent();
    event.id = simple.id;
    event.title = simple.title || "";
    event.descriptionText = buildDescriptionText(simple);
    if (simple.location) {
      event.setProperty("LOCATION", simple.location);
    }
    event.startDate = cal.dtz.fromRFC3339(simple.startISO, cal.dtz.UTC);
    event.endDate = cal.dtz.fromRFC3339(simple.endISO, cal.dtz.UTC);
    if (simple.isAllDay) {
      event.startDate.isDate = true;
      event.endDate.isDate = true;
    }
    if (simple.organizer) {
      const organizer = new CalAttendee();
      organizer.id = `mailto:${simple.organizer.email}`;
      organizer.commonName = simple.organizer.name;
      organizer.isOrganizer = true;
      event.organizer = organizer;
    }

    for (const a of simple.attendees || []) {
      if (!a.email) {
        continue;
      }
      const attendee = new CalAttendee();
      attendee.id = `mailto:${a.email}`;
      attendee.commonName = a.name;
      attendee.role = a.role || "REQ-PARTICIPANT";
      attendee.participationStatus = a.status || "NEEDS-ACTION";
      event.addAttendee(attendee);
    }
    if (simple.recurrenceRule) {
      try {
        const recInfo = new CalRecurrenceInfo(event);
        const rule = new CalRecurrenceRule();
        rule.icalString = `RRULE:${simple.recurrenceRule}`;
        recInfo.appendRecurrenceItem(rule);
        event.recurrenceInfo = recInfo;
      } catch (e) {
        console.error("exchangeCalendar: could not apply recurrence rule", simple.recurrenceRule, e);
      }
    }

    // Cancelled meetings: STATUS=CANCELLED gets a native strikethrough in
    // Thunderbird's calendar views with zero configuration. The category
    // is additionally there so the user can optionally assign it a color
    // (e.g. red) in Calendar > Categories, since per-event colors in
    // Thunderbird are driven by categories, not a direct color property.
    const categories = [...(simple.categories || [])];
    if (simple.isCancelled) {
      event.status = "CANCELLED";
      if (!categories.includes(CANCELLED_CATEGORY)) {
        categories.push(CANCELLED_CATEGORY);
      }
    } else {
      event.status = "CONFIRMED";
    }
    event.setCategories(categories);

    // Thunderbird has no native 4-state (Free/Tentative/Busy/OOF) display
    // the way Outlook does, but setting TRANSP at least keeps "Free" time
    // from counting as busy for anything that queries this calendar's
    // free/busy.
    event.setProperty("TRANSP", simple.freeBusyStatus === "Free" ? "TRANSPARENT" : "OPAQUE");

    if (simple.reminderMinutesBeforeStart != null) {
      event.addAlarm(buildReminderAlarm(simple.reminderMinutesBeforeStart));
    }

    // EWS's concurrency token — required on every UpdateItem/DeleteItem
    // call. Round-tripped via a custom property since calIEvent has no
    // native concept of it; see calEventToSimple() below.
    if (simple.changeKey) {
      event.setProperty("X-EWS-CHANGEKEY", simple.changeKey);
    }
    return event;
  }

  /**
   * Thunderbird shows Accept / Tentative / Decline only when it can find
   * "the invited attendee": either the property X-MOZ-INVITED-ATTENDEE on
   * the item, or calendar.organizerId matched against an attendee who is
   * not the organizer. Stamp both.
   */
  function stampInvitation(calendar, event, simple) {
    const organizerId = calendar.getProperty("organizerId");
    if (!organizerId || !event.organizer) {
      return;
    }
    const owner = organizerId.toLowerCase();
    if ((event.organizer.id || "").toLowerCase() === owner) {
      return;
    }
    const status = simple?.myResponseType ? ewsResponseToPartStat(simple.myResponseType) : null;
    const attendees = event.getAttendees();
    let match = null;
    event.removeAllAttendees();
    for (const att of attendees) {
      if ((att.id || "").toLowerCase() === owner) {
        if (status) {
          att.participationStatus = status;
        }
        match = att;
      }
      event.addAttendee(att);
    }
    if (match) {
      event.setProperty("X-MOZ-INVITED-ATTENDEE", match.id);
    }
  }

  function adoptForCalendar(calendar, item) {
    try {
      item.calendar = calendar;
    } catch (e) {
      /* stored items are sometimes immutable; the property stamp still works */
    }
    return item;
  }

  function calEventToSimple(event) {
    const toCleanISO = (calDate) => {
      if (!calDate) return null;
      const utc = calDate.getInTimezone(cal.dtz.UTC);
      const js = cal.dtz.dateTimeToJsDate(utc);
      return js.toISOString().replace(/\.\d{3}Z$/, "Z");
    };

    return {
      id: event.id,
      changeKey: event.getProperty("X-EWS-CHANGEKEY") || null,
      title: event.title || "",
      description: event.descriptionText || "",
      location: event.getProperty("LOCATION") || null,
      startISO: toCleanISO(event.startDate),
      endISO: toCleanISO(event.endDate),
      isAllDay: !!event.startDate.isDate,
      organizer: event.organizer
        ? { name: event.organizer.commonName, email: (event.organizer.id || "").replace(/^mailto:/i, "") }
        : null,
      attendees: event.getAttendees().map(a => ({
        name: a.commonName,
        email: (a.id || "").replace(/^mailto:/i, ""),
        role: a.role === "OPT-PARTICIPANT" ? "OPT-PARTICIPANT" : "REQ-PARTICIPANT",
        status: a.participationStatus || "NEEDS-ACTION",
      })),
      // Our own synthetic "Cancelled" tag (see simpleToCalEvent) isn't a
      // real Outlook category, so it's never pushed back.
      categories: event.getCategories().filter(c => c !== CANCELLED_CATEGORY),
      reminderMinutesBeforeStart: getReminderMinutesBeforeStart(event),
      // TODO: round-trip recurrenceInfo back to an RRULE string for edits
      // to recurring events (see ews.js simpleToEwsItemXml for the
      // matching TODO on the pull side).
      recurrenceRule: null,
    };
  }

  class ExchangeEwsCalendar extends cal.provider.BaseClass {
    constructor() {
      super();
      this.initProviderBase();
      // Bound lazily in the `id` setter below — see bindStore().
      this.store = null;
    }

    // Track which calendarKey this calendar is, and bind its backing
    // store, as soon as Thunderbird assigns it a (persistent) id — both on
    // first registration and when Thunderbird recreates it from prefs on a
    // later startup. This can't happen in the constructor: until our own
    // id is set, getProperty()/setProperty() have nothing to persist to
    // (see cal.provider.BaseClass), so we couldn't remember the store's id
    // across restarts.
    get id() {
      return super.id;
    }
    set id(value) {
      super.id = value;
      const calendarKey = this.getProperty("exchangeCalendarKey");
      if (calendarKey) {
        calendarsByKey.set(calendarKey, this);
      }
      this.bindStore();
    }

    // Backs this calendar with a real, persistent "storage" (SQLite)
    // calendar instead of the ephemeral "memory" type that earlier
    // versions used — "memory" never writes to disk, so every synced
    // event vanished on every Thunderbird restart even though the calendar
    // registration itself (being prefs-backed) survived. This reuses the
    // same shared moz-storage-calendar database every local/offline
    // calendar in the profile uses, under a stable id of our own so we
    // find the same rows again on the next startup.
    bindStore() {
      if (this.store) {
        this.wasStoreReset = false;
        return;
      }
      let storeId = this.getProperty("storeCalendarId");
      // No storeId yet means either a brand-new calendar, or one upgrading
      // from a version that used the ephemeral "memory" backing (which had
      // no such property) — either way, whatever's in background.js's
      // saved sync state no longer matches what's actually on disk here,
      // so registerCalendar() reports this back for it to force a full
      // resync rather than a delta one.
      this.wasStoreReset = !storeId;
      if (!storeId) {
        storeId = cal.getUUID();
        this.setProperty("storeCalendarId", storeId);
      }
      const store = Cc["@mozilla.org/calendar/calendar;1?type=storage"].createInstance(Ci.calICalendar);
      store.superCalendar = this;
      store.uri = Services.io.newURI("moz-storage-calendar://");
      store.id = storeId;
      store.addObserver(new RelayObserver(this));
      this.store = store;
    }

    get type() {
      return CALENDAR_TYPE;
    }

    get canRefresh() {
      // We're push/pull-synced by background.js on its own alarm schedule,
      // not by Thunderbird calling refresh(); see README.
      return false;
    }

    // Thunderbird's invitation bar calls supportsScheduling, then
    // getSchedulingSupport().getInvitedAttendee(). Without this the
    // Accept / Tentative / Decline row stays hidden.
    get supportsScheduling() {
      return !!this.getProperty("organizerId");
    }
    getSchedulingSupport() {
      return this;
    }
    getInvitedAttendee(aItem) {
      const preset = aItem.getProperty("X-MOZ-INVITED-ATTENDEE");
      if (preset) {
        const found = aItem.getAttendeeById(preset);
        if (found) {
          return found;
        }
      }
      const id = (this.getProperty("organizerId") || "").toLowerCase();
      if (!id || !aItem.organizer) {
        return null;
      }
      if ((aItem.organizer.id || "").toLowerCase() === id) {
        return null;
      }
      for (const att of aItem.getAttendees()) {
        if ((att.id || "").toLowerCase() === id) {
          return att;
        }
      }
      return null;
    }
    canNotify() {
      // The Exchange server sends the meeting response. Returning true
      // stops Thunderbird from also trying to send its own iTIP email.
      return true;
    }

    // Thunderbird 128+ expects a ReadableStream<calIItemBase> here (not a
    // Promise of an array), so hand the store's stream straight through.
    getItems(itemFilter, count, rangeStart, rangeEnd) {
      return this.store.getItems(itemFilter, count, rangeStart, rangeEnd);
    }

    async getItem(id) {
      const item = await this.store.getItem(id);
      return item ? adoptForCalendar(this, item) : null;
    }

    async applyRemoteChange(change) {
      const { item } = change;
      const existing = await this.store.getItem(item.id);
      if (change.op === "delete" || item.removed) {
        if (existing) {
          await this.store.deleteItem(existing);
        }
        return;
      }
      const calEvent = simpleToCalEvent(item);
      stampInvitation(this, calEvent, item);
      adoptForCalendar(this, calEvent);
      if (existing) {
        await this.store.modifyItem(calEvent, existing);
      } else {
        await this.store.addItem(calEvent);
      }
    }

    async addItem(item) {
      return this.pushLocalChange("add", item, null);
    }

    async modifyItem(newItem, oldItem) {
      return this.pushLocalChange("modify", newItem, oldItem);
    }

    async deleteItem(item) {
      return this.pushLocalChange("delete", item, null);
    }

    async pushLocalChange(op, item, oldItem) {
      if (!fireLocalChange) {
        throw new Components.Exception(
          "Exchange Calendar Sync's background page isn't running",
          Cr.NS_ERROR_NOT_AVAILABLE
        );
      }
      const calendarKey = this.getProperty("exchangeCalendarKey");
      const requestId = cal.getUUID();
      const resultPromise = new Promise((resolve, reject) => {
        pendingLocalChanges.set(requestId, { resolve, reject });
      });
      fireLocalChange(calendarKey, op, calEventToSimple(item), oldItem ? calEventToSimple(oldItem) : null, requestId);

      let resultSimple;
      try {
        resultSimple = await resultPromise;
      } finally {
        pendingLocalChanges.delete(requestId);
      }

      if (op === "delete") {
        const existing = await this.store.getItem(item.id);
        if (existing) {
          await this.store.deleteItem(existing);
        }
        return null;
      }

      // For adds, the original item usually has a temporary local id.
      // After the server assigns a real ItemId we must remove the temporary
      // one so the view doesn't show a duplicate or a broken placeholder.
      if (op === "add" && item.id && resultSimple?.id && item.id !== resultSimple.id) {
        try {
          const temp = await this.store.getItem(item.id);
          if (temp) {
            await this.store.deleteItem(temp);
          }
        } catch (e) {
          console.warn("exchangeCalendar: could not remove temporary local item", e);
        }
      }

      const calEvent = simpleToCalEvent(resultSimple);
      stampInvitation(this, calEvent, resultSimple);
      adoptForCalendar(this, calEvent);
      const existing = await this.store.getItem(calEvent.id);
      if (existing) {
        return this.store.modifyItem(calEvent, existing);
      }
      return this.store.addItem(calEvent);
    }
  }

  /** Relays the inner store's notifications up to this (outer, registered) calendar's own observers. */
  class RelayObserver {
    QueryInterface = ChromeUtils.generateQI(["calIObserver"]);

    constructor(calendar) {
      this.calendar = calendar;
    }
    onStartBatch() {
      this.calendar.observers.notify("onStartBatch", [this.calendar]);
    }
    onEndBatch() {
      this.calendar.observers.notify("onEndBatch", [this.calendar]);
    }
    onLoad() {
      this.calendar.observers.notify("onLoad", [this.calendar]);
    }
    onAddItem(item) {
      this.calendar.observers.notify("onAddItem", [item]);
    }
    onModifyItem(newItem, oldItem) {
      this.calendar.observers.notify("onModifyItem", [newItem, oldItem]);
    }
    onDeleteItem(item) {
      this.calendar.observers.notify("onDeleteItem", [item]);
    }
    onError(calendar, errNo, message) {
      this.calendar.observers.notify("onError", [this.calendar, errNo, message]);
    }
    onPropertyChanged(calendar, name, value, oldValue) {
      this.calendar.observers.notify("onPropertyChanged", [this.calendar, name, value, oldValue]);
    }
    onPropertyDeleting(calendar, name) {
      this.calendar.observers.notify("onPropertyDeleting", [this.calendar, name]);
    }
  }

  if (!cal.manager.hasCalendarProvider(CALENDAR_TYPE)) {
    cal.manager.registerCalendarProvider(CALENDAR_TYPE, ExchangeEwsCalendar);
  }

  class ExchangeCalendarAPI extends ExtensionCommon.ExtensionAPI {
    onShutdown(isAppShutdown) {
      if (isAppShutdown) {
        return;
      }
      try {
        cal.manager.unregisterCalendarProvider(CALENDAR_TYPE, true);
      } catch (e) {
        console.error("exchangeCalendar: unregister on shutdown failed", e);
      }
    }

    getAPI(context) {
      return {
        exchangeCalendar: {
          async registerCalendar(calendarKey, displayName, ownerEmail) {
            const applyOwner = calendar => {
              if (!ownerEmail || !ownerEmail.includes("@")) {
                return;
              }
              const mailto = ownerEmail.toLowerCase().startsWith("mailto:")
                ? ownerEmail
                : `mailto:${ownerEmail}`;
              calendar.setProperty("organizerId", mailto);
              calendar.setProperty("organizerCN", ownerEmail);
              calendar.setProperty("imip.identity.disabled", false);
            };
            let calendar = calendarsByKey.get(calendarKey);
            if (calendar) {
              calendar.name = displayName;
              applyOwner(calendar);
              return { calendarId: calendar.id, wasStoreReset: !!calendar.wasStoreReset };
            }
            calendar = new ExchangeEwsCalendar();
            calendar.setProperty("exchangeCalendarKey", calendarKey);
            applyOwner(calendar);
            calendar.name = displayName;
            calendar.uri = Services.io.newURI(`exchangeewssync://${encodeURIComponent(calendarKey)}/`);
            cal.manager.registerCalendar(calendar);
            calendarsByKey.set(calendarKey, calendar);
            return { calendarId: calendar.id, wasStoreReset: !!calendar.wasStoreReset };
          },

          async unregisterCalendar(calendarKey) {
            const calendar = calendarsByKey.get(calendarKey);
            if (calendar) {
              cal.manager.unregisterCalendar(calendar);
              calendarsByKey.delete(calendarKey);
            }
          },

          // Credentials go through nsILoginManager (the same encrypted
          // store IMAP/SMTP passwords use — protected by the user's Primary
          // Password if they've set one) rather than browser.storage.local,
          // since this is a long-lived reusable domain password rather than
          // a short-lived OAuth token. Keyed by accountId (the EWS login),
          // not calendarKey — several synced calendars can share one login.
          async saveCredentials(accountId, username, password) {
            const { origin, httpRealm } = credentialScope(accountId);
            await clearStoredCredentials(accountId);
            const loginInfo = Cc["@mozilla.org/login-manager/loginInfo;1"].createInstance(Ci.nsILoginInfo);
            loginInfo.init(origin, null, httpRealm, username, password, "", "");
            await Services.logins.addLoginAsync(loginInfo);
          },

          async getCredentials(accountId) {
            const { origin, httpRealm } = credentialScope(accountId);
            const [login] = await Services.logins.searchLoginsAsync({ origin, httpRealm });
            return login ? { username: login.username, password: login.password } : null;
          },

          async deleteCredentials(accountId) {
            await clearStoredCredentials(accountId);
          },

          async applyRemoteChanges(calendarKey, changes) {
            const calendar = calendarsByKey.get(calendarKey);
            if (!calendar) {
              throw new Error(`exchangeCalendar: no registered calendar for ${calendarKey}`);
            }
            // One bad item must not abort the rest of the batch (it used to,
            // so every item after it silently never synced). Failures are
            // returned so background.js can retry them on the next sync.
            const failedIds = [];
            calendar.startBatch();
            try {
              for (const change of changes) {
                try {
                  await calendar.applyRemoteChange(change);
                } catch (e) {
                  console.error(`exchangeCalendar: could not apply ${change.op} for ${change.item?.id}`, e);
                  failedIds.push(change.item?.id);
                }
              }
            } finally {
              calendar.endBatch();
            }
            return failedIds;
          },

          async resolveLocalChange(requestId, resultItem) {
            const pending = pendingLocalChanges.get(requestId);
            if (pending) {
              pending.resolve(resultItem);
            }
          },

          async rejectLocalChange(requestId, message) {
            const pending = pendingLocalChanges.get(requestId);
            if (pending) {
              pending.reject(new Error(message));
            }
          },

          onLocalChange: new ExtensionCommon.EventManager({
            context,
            name: "exchangeCalendar.onLocalChange",
            register(fire) {
              fireLocalChange = (...args) => fire.async(...args);
              return () => {
                fireLocalChange = null;
              };
            },
          }).api(),
        },
      };
    }
  }

  exports.exchangeCalendar = ExchangeCalendarAPI;
})(this);
