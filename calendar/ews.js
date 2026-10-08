/*
 * ews.js — Exchange Web Services (SOAP) client for an on-premises Exchange
 * server.
 *
 * Runs in the regular (unprivileged) WebExtension background page, loaded
 * as a plain script before background.js (see manifest.json). Only uses
 * fetch() and DOMParser — both ordinary web APIs available to any
 * WebExtension background page, no XPCOM here.
 *
 * Authentication: requests are sent with no Authorization header at all.
 * The server's resulting 401 challenge (Basic, NTLM, or Negotiate —
 * whichever it wants) is answered by background.js's
 * `browser.webRequest.onAuthRequired` listener, which hands Gecko's own
 * network stack the username/password; Gecko completes the handshake.
 * This file never sets that header itself — doing so would make Gecko
 * treat auth as already handled and skip that listener entirely.
 *
 * Every FieldURI/operation shape here was verified against Microsoft's own
 * EWS Managed API source (OfficeDev/ews-managed-api on GitHub) by grepping
 * the real files directly, not from memory — see the comments at each
 * call site for what was checked. In particular, AcceptItem/DeclineItem/
 * TentativelyAcceptItem are NOT separate top-level SOAP operations; they're
 * item types created via the same CreateItem operation used for new
 * events (confirmed via CreateResponseObjectRequest, which extends
 * CreateItemRequestBase).
 *
 * Exposes:
 *   - class EwsClient        SOAP calls: listCalendarFolders /
 *                             syncFolderItems / createEvent / updateEvent /
 *                             deleteEvent / respondToInvite
 *   - ewsItemToSimple(xmlEl) / simpleToEwsItemXml(simpleEvent)
 *   - PRIMARY_CALENDAR       the folderRef for the account's default calendar
 */

const SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/";
const TYPES_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
const MSGS_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
// A version tag old enough to be accepted by Exchange 2013 SP1 through 2019
// (EWS rejects a tag *newer* than the mailbox's server version, but happily
// accepts an older one — so this is the safe common denominator rather than
// the newest schema).
const SERVER_VERSION = "Exchange2013_SP1";

// A folderRef is either { distinguishedId: "calendar" } (the account's
// primary calendar) or { id: "<EWS folder id>" } (a secondary calendar,
// discovered via listCalendarFolders()).
const PRIMARY_CALENDAR = { distinguishedId: "calendar" };

function folderRefXml(folderRef) {
  return folderRef.distinguishedId
    ? `<t:DistinguishedFolderId Id="${escapeXmlAttr(folderRef.distinguishedId)}"/>`
    : `<t:FolderId Id="${escapeXmlAttr(folderRef.id)}"/>`;
}

class EwsSoapFault extends Error {
  constructor(message, responseCode) {
    super(message);
    this.name = "EwsSoapFault";
    this.responseCode = responseCode;
  }
}

/**
 * Turns an HTTP-level SOAP <Fault> body (distinct from the ResponseClass
 * faults assertSuccess() handles — this is for failures before EWS even
 * gets to process the operation, like schema validation) into a readable
 * message, pulling faultcode/faultstring and whatever the <detail>
 * element holds rather than a blind text slice that can truncate the
 * one part (the detail) that says *what* was invalid.
 */
function describeSoapFault(text) {
  const doc = new DOMParser().parseFromString(text, "text/xml");
  const fault = doc.getElementsByTagNameNS(SOAP_NS, "Fault")[0];
  if (!fault) {
    return text.slice(0, 2000);
  }
  const code = fault.getElementsByTagName("faultcode")[0]?.textContent;
  const message = fault.getElementsByTagName("faultstring")[0]?.textContent;
  const detail = fault.getElementsByTagName("detail")[0];
  const detailText = detail ? new XMLSerializer().serializeToString(detail) : "";
  return [code, message, detailText].filter(Boolean).join(" — ");
}

class EwsClient {
  constructor(url, username, password) {
    this.url = url;
    this.username = username;
    this.password = password;
  }

  async soapRequest(bodyXml) {
    const envelope =
      `<?xml version="1.0" encoding="utf-8"?>` +
      `<soap:Envelope xmlns:soap="${SOAP_NS}" xmlns:t="${TYPES_NS}" xmlns:m="${MSGS_NS}">` +
      // No TimeZoneContext header: CalendarItem Start/End come back as UTC
      // by default anyway, and getting that header's schema wrong is a
      // likely source of "ErrorInvalidRequest" schema-validation faults.
      `<soap:Header><t:RequestServerVersion Version="${SERVER_VERSION}"/></soap:Header>` +
      `<soap:Body>${bodyXml}</soap:Body>` +
      `</soap:Envelope>`;

    // Deliberately no Authorization header here: setting one ourselves
    // would make Gecko treat auth as "already handled by the caller" and
    // skip its own 401-challenge handling, which is what actually drives
    // background.js's webRequest.onAuthRequired listener. Leaving this
    // request unauthenticated lets the server's 401 challenge (Basic,
    // NTLM, or Negotiate — whichever it wants) reach that listener, which
    // answers with the stored username/password regardless of scheme.
    const resp = await fetch(this.url, {
      method: "POST",
      headers: { "Content-Type": 'text/xml; charset="utf-8"' },
      body: envelope,
    });
    const text = await resp.text();
    if (!resp.ok) {
      const authHeader = resp.headers.get("WWW-Authenticate");
      throw new Error(
        `EWS HTTP ${resp.status} for ${this.url}` +
          (authHeader ? ` (server wants: ${authHeader})` : "") +
          `: ${describeSoapFault(text)}`
      );
    }
    const doc = new DOMParser().parseFromString(text, "text/xml");
    if (doc.querySelector("parsererror")) {
      throw new Error(`EWS response was not valid XML: ${text.slice(0, 2000)}`);
    }
    return doc;
  }

  assertSuccess(doc, opName) {
    const responseMessages = [...doc.getElementsByTagName("*")].filter(
      el => el.getAttribute("ResponseClass") != null
    );
    for (const msg of responseMessages) {
      const responseClass = msg.getAttribute("ResponseClass");
      if (responseClass !== "Success") {
        const codeEl = msg.getElementsByTagNameNS(MSGS_NS, "ResponseCode")[0];
        const textEl = msg.getElementsByTagNameNS(MSGS_NS, "MessageText")[0];
        throw new EwsSoapFault(
          `${opName} failed: ${textEl?.textContent || responseClass}`,
          codeEl?.textContent
        );
      }
    }
    return responseMessages;
  }

  /**
   * Lists calendars available on this account: the primary calendar plus
   * any secondary ones (e.g. "Team Events") created as child folders of
   * it — which is where Outlook puts them when you use "Add Calendar" ->
   * "New Blank Calendar". Does NOT find calendars shared/delegated from a
   * different mailbox; that's a separate, unimplemented feature.
   *
   * FindFolder/FolderShape/DisplayName/FolderClass FieldURIs verified
   * against FindFolderRequest.cs, FindRequest.cs, FolderView.cs, and
   * FolderSchema.cs in OfficeDev/ews-managed-api.
   *
   * @returns {{folderRef: object, name: string}[]}
   */
  async listCalendarFolders() {
    const body =
      `<m:FindFolder Traversal="Shallow">` +
      `<m:FolderShape><t:BaseShape>IdOnly</t:BaseShape>` +
      `<t:AdditionalProperties>` +
      `<t:FieldURI FieldURI="folder:DisplayName"/>` +
      `<t:FieldURI FieldURI="folder:FolderClass"/>` +
      `</t:AdditionalProperties></m:FolderShape>` +
      `<m:ParentFolderIds>${folderRefXml(PRIMARY_CALENDAR)}</m:ParentFolderIds>` +
      `</m:FindFolder>`;
    const doc = await this.soapRequest(body);
    this.assertSuccess(doc, "FindFolder");

    const folders = [{ folderRef: PRIMARY_CALENDAR, name: "Calendar" }];
    for (const folderEl of doc.getElementsByTagNameNS(TYPES_NS, "CalendarFolder")) {
      const folderClass = text(folderEl, "FolderClass");
      if (folderClass && folderClass !== "IPF.Appointment") {
        continue;
      }
      const folderId = folderEl.getElementsByTagNameNS(TYPES_NS, "FolderId")[0];
      const name = text(folderEl, "DisplayName");
      if (folderId && name) {
        folders.push({ folderRef: { id: folderId.getAttribute("Id") }, name });
      }
    }
    return folders;
  }

  /**
   * Incrementally syncs a calendar folder. Pass the syncState saved from a
   * previous call to get only what changed since; omit it for an initial
   * full sync. Call repeatedly while `moreAvailable` is true.
   *
   * @returns {{changes: object[], syncState: string, moreAvailable: boolean}}
   */
  async syncFolderItems(folderRef, syncState) {
    const body =
      `<m:SyncFolderItems>` +
      `<m:ItemShape><t:BaseShape>Default</t:BaseShape>` +
      `<t:AdditionalProperties>` +
      `<t:FieldURI FieldURI="calendar:Start"/>` +
      `<t:FieldURI FieldURI="calendar:End"/>` +
      `<t:FieldURI FieldURI="calendar:Location"/>` +
      `<t:FieldURI FieldURI="calendar:IsAllDayEvent"/>` +
      `<t:FieldURI FieldURI="calendar:IsCancelled"/>` +
      `<t:FieldURI FieldURI="item:ReminderMinutesBeforeStart"/>` +
      `<t:FieldURI FieldURI="item:ReminderIsSet"/>` +
      `<t:FieldURI FieldURI="item:Categories"/>` +
      `<t:FieldURI FieldURI="calendar:Organizer"/>` +
      `<t:FieldURI FieldURI="calendar:RequiredAttendees"/>` +
      `<t:FieldURI FieldURI="calendar:OptionalAttendees"/>` +
      `<t:FieldURI FieldURI="calendar:UID"/>` +
      `<t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus"/>` +
      `<t:FieldURI FieldURI="calendar:MyResponseType"/>` +
      `<t:FieldURI FieldURI="calendar:IsOnlineMeeting"/>` +
      `<t:FieldURI FieldURI="calendar:JoinOnlineMeetingUrl"/>` +
      `</t:AdditionalProperties></m:ItemShape>` +
      `<m:SyncFolderId>${folderRefXml(folderRef)}</m:SyncFolderId>` +
      (syncState ? `<m:SyncState>${escapeXml(syncState)}</m:SyncState>` : "") +
      `<m:MaxChangesReturned>200</m:MaxChangesReturned>` +
      `</m:SyncFolderItems>`;

    const doc = await this.soapRequest(body);
    this.assertSuccess(doc, "SyncFolderItems");

    const changes = [];
    for (const created of doc.getElementsByTagNameNS(TYPES_NS, "Create")) {
      const item = created.getElementsByTagNameNS(TYPES_NS, "CalendarItem")[0];
      if (item) {
        changes.push({ op: "create", item: ewsItemToSimple(item) });
      }
    }
    for (const updated of doc.getElementsByTagNameNS(TYPES_NS, "Update")) {
      const item = updated.getElementsByTagNameNS(TYPES_NS, "CalendarItem")[0];
      if (item) {
        changes.push({ op: "update", item: ewsItemToSimple(item) });
      }
    }
    for (const deleted of doc.getElementsByTagNameNS(TYPES_NS, "Delete")) {
      const itemId = deleted.getElementsByTagNameNS(TYPES_NS, "ItemId")[0];
      if (itemId) {
        changes.push({ op: "delete", item: { id: itemId.getAttribute("Id"), removed: true } });
      }
    }

    const syncStateEl = doc.getElementsByTagNameNS(MSGS_NS, "SyncState")[0];
    const moreEl = doc.getElementsByTagNameNS(MSGS_NS, "IncludesLastItemInRange")[0];
    return {
      changes,
      syncState: syncStateEl?.textContent || syncState,
      moreAvailable: moreEl ? moreEl.textContent !== "true" : false,
    };
  }

  async createEvent(folderRef, simpleEvent) {
    const body =
      `<m:CreateItem SendMeetingInvitations="SendToNone">` +
      `<m:SavedItemFolderId>${folderRefXml(folderRef)}</m:SavedItemFolderId>` +
      `<m:Items>${simpleToEwsItemXml(simpleEvent)}</m:Items>` +
      `</m:CreateItem>`;
    const doc = await this.soapRequest(body);
    this.assertSuccess(doc, "CreateItem");
    const item = doc.getElementsByTagNameNS(TYPES_NS, "CalendarItem")[0];
    return ewsItemToSimple(item);
  }

  async updateEvent(simpleEvent) {
    const fields = [
      ["item:Subject", `<t:Subject>${escapeXml(simpleEvent.title || "")}</t:Subject>`],
      ["item:Body", `<t:Body BodyType="Text">${escapeXml(simpleEvent.description || "")}</t:Body>`],
      ["calendar:Start", `<t:Start>${simpleEvent.startISO}</t:Start>`],
      ["calendar:End", `<t:End>${simpleEvent.endISO}</t:End>`],
      ["calendar:Location", `<t:Location>${escapeXml(simpleEvent.location || "")}</t:Location>`],
      ["calendar:IsAllDayEvent", `<t:IsAllDayEvent>${!!simpleEvent.isAllDay}</t:IsAllDayEvent>`],
      ["item:Categories", categoriesXml(simpleEvent.categories)],
    ];
    const setFields = fields
      .map(
        ([fieldUri, xml]) =>
          `<t:SetItemField><t:FieldURI FieldURI="${fieldUri}"/><t:CalendarItem>${xml}</t:CalendarItem></t:SetItemField>`
      )
      .join("");

    const body =
      `<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">` +
      `<m:ItemChanges><t:ItemChange>` +
      `<t:ItemId Id="${escapeXmlAttr(simpleEvent.id)}" ChangeKey="${escapeXmlAttr(simpleEvent.changeKey || "")}"/>` +
      `<t:Updates>${setFields}</t:Updates>` +
      `</t:ItemChange></m:ItemChanges>` +
      `</m:UpdateItem>`;
    const doc = await this.soapRequest(body);
    this.assertSuccess(doc, "UpdateItem");
    const item = doc.getElementsByTagNameNS(TYPES_NS, "CalendarItem")[0];
    return item ? ewsItemToSimple(item) : simpleEvent;
  }

  async deleteEvent(simpleEvent) {
    const body =
      `<m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToNone">` +
      `<m:ItemIds><t:ItemId Id="${escapeXmlAttr(simpleEvent.id)}" ChangeKey="${escapeXmlAttr(simpleEvent.changeKey || "")}"/></m:ItemIds>` +
      `</m:DeleteItem>`;
    const doc = await this.soapRequest(body);
    this.assertSuccess(doc, "DeleteItem");
  }

  /**
   * Accepts, declines, or tentatively accepts a meeting invitation,
   * sending the response to the organizer and updating this item in our
   * own calendar to match.
   *
   * AcceptItem/DeclineItem/TentativelyAcceptItem are not separate EWS
   * operations — they're item types submitted through the same CreateItem
   * operation used to create a new event, which is why this reuses the
   * CreateItem shape rather than looking like createEvent()'s sibling.
   * Verified via CreateResponseObjectRequest (extends
   * CreateItemRequestBase) and ResponseObjectSchema.ReferenceItemId in
   * ews-managed-api.
   *
   * @param {"ACCEPTED"|"DECLINED"|"TENTATIVE"} response
   */
  async respondToInvite(simpleEvent, response) {
    const elementName = { ACCEPTED: "AcceptItem", DECLINED: "DeclineItem", TENTATIVE: "TentativelyAcceptItem" }[
      response
    ];
    const body =
      `<m:CreateItem MessageDisposition="SendAndSaveCopy">` +
      `<m:Items><t:${elementName}>` +
      `<t:ReferenceItemId Id="${escapeXmlAttr(simpleEvent.id)}" ChangeKey="${escapeXmlAttr(simpleEvent.changeKey || "")}"/>` +
      `</t:${elementName}></m:Items>` +
      `</m:CreateItem>`;
    const doc = await this.soapRequest(body);
    this.assertSuccess(doc, "CreateItem (meeting response)");
    const item = doc.getElementsByTagNameNS(TYPES_NS, "CalendarItem")[0];
    return item ? ewsItemToSimple(item) : simpleEvent;
  }
}

function text(el, localName) {
  return el.getElementsByTagNameNS(TYPES_NS, localName)[0]?.textContent || null;
}

function mailboxToSimple(mailboxEl) {
  if (!mailboxEl) {
    return null;
  }
  return {
    name: text(mailboxEl, "Name") || text(mailboxEl, "EmailAddress"),
    email: text(mailboxEl, "EmailAddress"),
  };
}

/** <t:Categories><t:String>A</t:String><t:String>B</t:String></t:Categories> — the standard EWS ArrayOfStringsType shape. */
function categoriesXml(categories) {
  const entries = (categories || []).map(c => `<t:String>${escapeXml(c)}</t:String>`).join("");
  return `<t:Categories>${entries}</t:Categories>`;
}

/** Converts a <t:CalendarItem> element from an EWS response into the plain "SimpleEvent" shape shared with provider.js. */
function ewsItemToSimple(item) {
  const itemId = item.getElementsByTagNameNS(TYPES_NS, "ItemId")[0];
  const organizerMailbox = item.getElementsByTagNameNS(TYPES_NS, "Organizer")[0]?.getElementsByTagNameNS(TYPES_NS, "Mailbox")[0];

  const attendees = [];
  for (const [tag, role] of [
    ["RequiredAttendees", "REQ-PARTICIPANT"],
    ["OptionalAttendees", "OPT-PARTICIPANT"],
  ]) {
    const container = item.getElementsByTagNameNS(TYPES_NS, tag)[0];
    if (!container) {
      continue;
    }
    for (const attendeeEl of container.getElementsByTagNameNS(TYPES_NS, "Attendee")) {
      const mailbox = attendeeEl.getElementsByTagNameNS(TYPES_NS, "Mailbox")[0];
      const responseType = text(attendeeEl, "ResponseType");
      attendees.push({
        name: text(mailbox, "Name") || text(mailbox, "EmailAddress"),
        email: text(mailbox, "EmailAddress"),
        role,
        status: ewsResponseToPartStat(responseType),
      });
    }
  }

  const categoriesContainer = item.getElementsByTagNameNS(TYPES_NS, "Categories")[0];
  const categories = categoriesContainer
    ? [...categoriesContainer.getElementsByTagNameNS(TYPES_NS, "String")].map(el => el.textContent)
    : [];

  const isReminderOn = text(item, "ReminderIsSet") === "true";

  return {
    id: itemId?.getAttribute("Id"),
    changeKey: itemId?.getAttribute("ChangeKey"),
    title: text(item, "Subject") || "",
    description: htmlToText(item.getElementsByTagNameNS(TYPES_NS, "Body")[0]?.textContent || ""),
    location: text(item, "Location"),
    startISO: text(item, "Start"),
    endISO: text(item, "End"),
    isAllDay: text(item, "IsAllDayEvent") === "true",
    isCancelled: text(item, "IsCancelled") === "true",
    organizer: mailboxToSimple(organizerMailbox),
    attendees,
    categories,
    freeBusyStatus: text(item, "LegacyFreeBusyStatus"), // "Free" | "Tentative" | "Busy" | "OOF" | "WorkingElsewhere" | "NoData"
    myResponseType: text(item, "MyResponseType"),
    isOnlineMeeting: text(item, "IsOnlineMeeting") === "true",
    onlineMeetingUrl: text(item, "JoinOnlineMeetingUrl"),
    reminderMinutesBeforeStart: isReminderOn ? parseInt(text(item, "ReminderMinutesBeforeStart") || "0", 10) : null,
    // TODO: recurrence — EWS exposes this via <t:Recurrence> (RelativeYearlyRecurrence,
    // AbsoluteMonthlyRecurrence, etc.) plus separate IsRecurring/CalendarItemType
    // fields, structurally different enough from Graph's `recurrence` object that
    // it needs its own mapping. Left out of v1; see README "Known limitations".
    recurrenceRule: null,
  };
}

/** Reverse of ewsItemToSimple(), wrapped in a <t:CalendarItem> for CreateItem. */
function simpleToEwsItemXml(simpleEvent) {
  const attendeesXml = (tag, role) => {
    const matching = (simpleEvent.attendees || []).filter(a => (a.role || "REQ-PARTICIPANT") === role);
    if (!matching.length) {
      return "";
    }
    const entries = matching
      .map(
        a =>
          `<t:Attendee><t:Mailbox><t:Name>${escapeXml(a.name || a.email)}</t:Name><t:EmailAddress>${escapeXml(a.email)}</t:EmailAddress></t:Mailbox></t:Attendee>`
      )
      .join("");
    return `<t:${tag}>${entries}</t:${tag}>`;
  };

  let reminderXml = "";
  if (simpleEvent.reminderMinutesBeforeStart != null) {
    reminderXml =
      `<t:ReminderIsSet>true</t:ReminderIsSet>` +
      `<t:ReminderMinutesBeforeStart>${simpleEvent.reminderMinutesBeforeStart}</t:ReminderMinutesBeforeStart>`;
  }

  return (
    `<t:CalendarItem>` +
    `<t:Subject>${escapeXml(simpleEvent.title || "")}</t:Subject>` +
    `<t:Body BodyType="Text">${escapeXml(simpleEvent.description || "")}</t:Body>` +
    (simpleEvent.location ? `<t:Location>${escapeXml(simpleEvent.location)}</t:Location>` : "") +
    `<t:Start>${simpleEvent.startISO}</t:Start>` +
    `<t:End>${simpleEvent.endISO}</t:End>` +
    `<t:IsAllDayEvent>${!!simpleEvent.isAllDay}</t:IsAllDayEvent>` +
    reminderXml +
    (simpleEvent.categories?.length ? categoriesXml(simpleEvent.categories) : "") +
    attendeesXml("RequiredAttendees", "REQ-PARTICIPANT") +
    attendeesXml("OptionalAttendees", "OPT-PARTICIPANT") +
    // NOTE: recurrence isn't round-tripped on push yet — see the TODO on
    // ewsItemToSimple() above. Creating/editing a recurring event from
    // Thunderbird currently saves as a single instance.
    `</t:CalendarItem>`
  );
}

function ewsResponseToPartStat(responseType) {
  switch (responseType) {
    case "Accept":
      return "ACCEPTED";
    case "Decline":
      return "DECLINED";
    case "Tentative":
      return "TENTATIVE";
    default:
      return "NEEDS-ACTION";
  }
}

function htmlToText(html) {
  if (!html) {
    return "";
  }
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .trim();
}

/** Escapes text for use between XML tags. */
function escapeXml(str) {
  return String(str ?? "").replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
}

/** Escapes text for use inside a double-quoted XML attribute. */
function escapeXmlAttr(str) {
  return escapeXml(str).replace(/"/g, "&quot;");
}
