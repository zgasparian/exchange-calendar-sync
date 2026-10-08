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
 * Exposes:
 *   - class EwsClient        SOAP calls: syncFolderItems / createEvent /
 *                             updateEvent / deleteEvent
 *   - ewsItemToSimple(xmlEl) / simpleToEwsItemXml(simpleEvent)
 */

const SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/";
const TYPES_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
const MSGS_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
// A version tag old enough to be accepted by Exchange 2013 SP1 through 2019
// (EWS rejects a tag *newer* than the mailbox's server version, but happily
// accepts an older one — so this is the safe common denominator rather than
// the newest schema).
const SERVER_VERSION = "Exchange2013_SP1";

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
   * Incrementally syncs the primary calendar folder. Pass the syncState
   * saved from a previous call to get only what changed since; omit it for
   * an initial full sync. Call repeatedly while `moreAvailable` is true.
   *
   * @returns {{changes: object[], syncState: string, moreAvailable: boolean}}
   */
  async syncFolderItems(syncState) {
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
      `<t:FieldURI FieldURI="calendar:Organizer"/>` +
      `<t:FieldURI FieldURI="calendar:RequiredAttendees"/>` +
      `<t:FieldURI FieldURI="calendar:OptionalAttendees"/>` +
      `<t:FieldURI FieldURI="calendar:UID"/>` +
      `</t:AdditionalProperties></m:ItemShape>` +
      `<m:SyncFolderId><t:DistinguishedFolderId Id="calendar"/></m:SyncFolderId>` +
      (syncState ? `<m:SyncState>${escapeXml(syncState)}</m:SyncState>` : "") +
      `<m:MaxChangesReturned>200</m:MaxChangesReturned>` +
      `</m:SyncFolderItems>`;

    const doc = await this.soapRequest(body);
    this.assertSuccess(doc, "SyncFolderItems");

    const changes = [];
    for (const created of doc.getElementsByTagNameNS(TYPES_NS, "Create")) {
      const item = created.getElementsByTagNameNS(TYPES_NS, "CalendarItem")[0];
      if (item) {
        changes.push({ op: "upsert", item: ewsItemToSimple(item) });
      }
    }
    for (const updated of doc.getElementsByTagNameNS(TYPES_NS, "Update")) {
      const item = updated.getElementsByTagNameNS(TYPES_NS, "CalendarItem")[0];
      if (item) {
        changes.push({ op: "upsert", item: ewsItemToSimple(item) });
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

  async createEvent(simpleEvent) {
    const body =
      `<m:CreateItem SendMeetingInvitations="SendToNone">` +
      `<m:SavedItemFolderId><t:DistinguishedFolderId Id="calendar"/></m:SavedItemFolderId>` +
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
