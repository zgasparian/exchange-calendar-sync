# Exchange Calendar Sync (Thunderbird extension)

Two-way sync between an **on-premises Exchange Server** mailbox's calendar
and Thunderbird, using EWS (Exchange Web Services) — no Microsoft 365 /
Azure AD / internet-facing identity provider involved. You give it a
server URL, a username, and a password, the same as you would in Outlook.

Developed by Zareh Kasparian, with Claude Code support.

## Provenance

Written from scratch — EWS against Microsoft's public
`learn.microsoft.com/exchange/client-developer/web-service-reference`
docs, and the Thunderbird integration against this machine's installed
Thunderbird 156 (`/snap/thunderbird/current/usr/lib/thunderbird/omni.ja`,
extracted and read directly: `CalCalendarManager.sys.mjs`,
`calProviderUtils.sys.mjs`, `CalDavCalendar.sys.mjs`, `CalMemoryCalendar.sys.mjs`,
`CalItemBase.sys.mjs`, `LoginManager.sys.mjs`/`CardDAVUtils.sys.mjs`, and
Thunderbird's own unfinished native `GraphCalendar.sys.mjs`/`GraphProvider.sys.mjs`
prototype). It does not contain or derive from any code from "Owl for
Exchange" (eule_fur_exchange), whose `LICENSE` file states its code is
proprietary and copying it is not permitted. Only its *folder layout* (an
Experiment API calendar provider plus a background page doing the network
talking) was used as a structural reference.

An earlier version of this extension used Microsoft Graph + OAuth2, which
only works against Exchange **Online** (Microsoft 365). That's gone now —
Graph has no on-premises equivalent, so this rewrite talks EWS instead.

## Features

- Two-way sync of events (create/edit/delete) between Thunderbird and an
  on-prem Exchange calendar.
- **Cancelled meetings** show with a strikethrough (native Thunderbird
  rendering for `STATUS=CANCELLED`) and a "Cancelled" category, so you can
  additionally assign that category a color (e.g. red) in Thunderbird's
  Calendar → Categories preferences if you want it to stand out more.
- **Reminders** sync in both directions.
- **Desktop notifications** for new meeting invites and for meetings that
  get cancelled (suppressed during a calendar's initial full sync, so
  connecting an account with months of history doesn't fire a notification
  storm).
- **Accept / Tentative / Decline** actually works: responding to an
  invitation from Thunderbird's own invitation UI sends a real meeting
  response to the organizer via EWS, not just a local-only status change.
- **Outlook categories** sync in both directions (`item:Categories`).
- **Free/busy**: an event's Exchange free/busy status maps to the
  iCalendar `TRANSP` property. Thunderbird doesn't have Outlook's 4-state
  (Free/Tentative/Busy/OOF) color-coded display, so this mainly matters
  for anything that queries this calendar's free/busy, not for how the
  event looks in your own view.
- **Online meeting links** (Teams, etc.) are appended to the event
  description when EWS reports one (`calendar:JoinOnlineMeetingUrl`).
- **Multiple calendars per account** — sync your primary calendar plus any
  secondary calendars (e.g. "Team Events") created under it. Shared or
  delegated calendars from a *different* mailbox aren't supported.
- Configurable sync interval, last-synced timestamps, encrypted credential
  storage — see below.

## Architecture

```
manifest.json            MV2, declares the exchangeCalendar Experiment API
background.js             Account/calendar storage, sync scheduling, pushes
                           local Thunderbird edits (including meeting
                           responses) to the server, answers HTTP auth
                           challenges (Basic/NTLM/Negotiate), desktop
                           notifications
calendar/ews.js            EWS SOAP client + EWS<->plain-JSON event mapping
                           (unprivileged — fetch() + DOMParser, no XPCOM)
calendar/schema.json       Experiment API definition (the privileged bridge)
calendar/provider.js       calICalendar implementation + Thunderbird
                           registration + encrypted credential storage
                           (privileged — the only file using XPCOM/ChromeUtils)
ui/options.html,options.js Settings page: connect/remove accounts, discover
                           and add/remove secondary calendars, sync interval
icons/                    Toolbar/about:addons icons
```

Networking lives entirely in the unprivileged background page
(`background.js` + `calendar/ews.js`), exactly like a normal WebExtension.
The privileged Experiment API (`calendar/provider.js`) only ever talks to
Thunderbird's calendar manager and to `nsILoginManager` — it registers one
`calICalendar` per *synced calendar* (backed internally by Thunderbird's
built-in "storage" calendar type — the same persistent SQLite-backed store
every local/offline calendar in the profile uses, so synced events survive
a Thunderbird restart), applies batches of changes background.js fetched
from EWS, relays the user's own edits back out via an `onLocalChange` event
so background.js can push them to the server, and stores/retrieves an
account's password encrypted-at-rest.

One EWS login (an "account") can own several synced calendars — its
primary one plus any secondary folders you add from the settings page.
Internally, `accountId` identifies the login (and is what credentials are
keyed by), while `calendarKey` (`${accountId}::${folderId}`) identifies
one specific synced calendar; `calendar/provider.js`'s module-level doc
comment spells out exactly which functions key off which.

(An earlier version used Thunderbird's "memory" calendar type instead,
which is never written to disk — every synced event vanished on every
restart even though the calendar's own registration survived. If you're
upgrading from that version, the first startup on the fixed version
detects the mismatch and forces one full resync automatically; no action
needed.)

### Upgrading from a pre-0.4.0 version

0.4.0 introduced the account/calendar split described above (one account,
multiple possible calendars) — before that, `accounts[accountId]` carried
its own sync state directly, since it was always exactly one calendar.
Upgrading in place handles this automatically: `background.js`'s
`migrateLegacyAccounts()` synthesizes the missing `calendars[calendarKey]`
entry for any such account (reusing its old sync token so this resumes
with a normal incremental sync, not a wasteful full one), and
`provider.js`'s `id` setter re-keys that account's *existing* Thunderbird
calendar under the new `calendarKey` instead of leaving it orphaned or
creating a duplicate alongside it. No action needed — but if you updated
from 0.4.0 itself (not earlier) and noticed new events silently stopped
syncing while old ones remained, that's exactly this bug; 0.4.1 fixes it.

## 1. Confirm EWS is reachable (IT-side check)

Ask whoever administers the Exchange server:

- **EWS must be enabled** for the mailbox (it is by default on Exchange
  2013–2019 unless an admin explicitly disabled it via
  `Set-CASMailbox -EwsEnabled $false`).
- You need the mail server's address — easiest is just the **OWA
  (webmail) address** you already use in a browser, e.g.
  `https://mail.yourcompany.com/owa`. The extension only keeps the
  hostname from it and builds the actual EWS endpoint
  (`https://<host>/EWS/Exchange.asmx` — a sibling of `/owa`, not nested
  under it) itself, since this extension doesn't implement Autodiscover.
  If OWA and EWS are ever split across *different* hostnames in your
  environment (uncommon, but possible with some reverse-proxy setups),
  ask IT for the EWS one specifically and enter that instead.
- The machine running Thunderbird needs **network access** to that URL —
  if Exchange is only reachable from inside the corporate network, you'll
  need to be on VPN or on-site.
- If the server uses a certificate from an **internal/private CA**
  (common for on-prem deployments), that CA needs to be trusted by
  Thunderbird/your OS already, or `fetch()` will fail with a TLS trust
  error — this extension can't and shouldn't bypass certificate validation.

## 2. Load the extension in Thunderbird

It's unsigned, so load it as a temporary add-on:

1. Thunderbird → hamburger menu → **Developer Tools** → **Debug Add-ons**
   (or go directly to `about:debugging`).
2. **This Thunderbird** → **Load Temporary Add-on…** → select this folder's
   `manifest.json`.
3. Open its **Preferences/Options** page (from about:addons, or it opens
   automatically).
4. Enter your **OWA/webmail address**, your **username** (whatever you'd type into
   Outlook's login prompt for this server — often `DOMAIN\username` or
   `username@company.com`, depending on how IIS auth is configured), and
   your **password**. Click **Connect account…**.

A new calendar (named after your username) should appear in Thunderbird's
calendar list within a few seconds, populated with your primary calendar's
events. Edits in either direction sync within 5 minutes by default (local
edits push immediately) via a `browser.alarms` timer in `background.js` —
the **Sync frequency** section on the settings page lets you change that
interval, and each synced calendar's row shows when it last synced
successfully.

To sync an additional (secondary) calendar on the same account, click
**Add calendar…** next to that account — this looks up every calendar
folder on the mailbox via EWS and lists whichever ones aren't already
synced. Each synced calendar gets its own entry (and its own "Remove"
button) in the list below its account.

### How authentication actually happens

Every EWS request goes out *without* an `Authorization` header. The server
responds 401 with a challenge for whatever scheme it wants — Basic, NTLM,
or Negotiate (Windows Integrated Auth, common for on-prem Exchange) —
and `background.js` answers that challenge via
`browser.webRequest.onAuthRequired`, handing Gecko's own network stack
your username/password; Gecko completes the actual handshake internally
and the final (authenticated) response is what `ews.js` sees. Either way,
you don't need to know which scheme your server uses — just enter your
normal login.

(Earlier builds set the `Authorization` header directly instead — don't
do that: it makes Gecko treat auth as "already handled by the caller"
and skip the challenge-response machinery `onAuthRequired` depends on,
so the listener never runs and a 401 passes straight through.)

This is also why the manifest asks for the broad `https://*/*` host
permission: the server's hostname isn't known until you type it into
settings, so the permission can't be scoped narrower ahead of time.

### Debugging

- `background.js` errors/logs show in the regular add-on **Inspect**
  console (same as any WebExtension). A connection failure (wrong URL,
  wrong credentials, EWS disabled, firewall) will show up here when you
  click **Connect account…**.
- `calendar/provider.js` runs in the privileged parent process — its errors
  show in Thunderbird's **Browser Console** (Ctrl+Shift+J), not the add-on
  inspector.

## Known limitations (v1)

- **Recurrence isn't implemented yet.** EWS exposes recurring events via a
  `<t:Recurrence>` element that's structurally different from a simple
  RRULE string, and needs its own mapping — left as a `TODO` in
  `calendar/ews.js` (`ewsItemToSimple`/`simpleToEwsItemXml`). Recurring
  events on the server currently won't sync correctly; non-recurring
  events are unaffected.
- **Free/busy has no dedicated UI.** Thunderbird doesn't render Outlook's
  4-state (Free/Tentative/Busy/OOF) color coding, so the
  `LegacyFreeBusyStatus` → `TRANSP` mapping is real but mostly invisible
  day-to-day; see "Features" above.
- **Shared/delegate calendars aren't supported** — only calendars that live
  directly on the account you authenticate as (primary + its own secondary
  folders). A calendar someone else has shared or delegated to you won't
  show up in "Add calendar…".
- **Credential storage**: the password is stored via `nsILoginManager`
  (`calendar/provider.js`'s `saveCredentials`/`getCredentials`), the same
  encrypted store Thunderbird uses for IMAP/SMTP passwords — protected by
  your Thunderbird Primary Password if you've set one. `background.js`
  only ever holds it in memory, never in `browser.storage.local`.
- **No real Autodiscover.** The EWS endpoint is derived from whatever
  hostname you enter (OWA address or otherwise) by assuming it's also the
  EWS host, which is true for the vast majority of on-prem setups — but
  isn't the full Autodiscover protocol, so an environment where OWA and
  EWS are deliberately split across hosts needs the EWS hostname entered
  directly instead.
- **Exchange Online / Microsoft 365** isn't supported by this version —
  it's EWS-only now. Cloud Exchange would need Microsoft Graph + OAuth2
  instead (a different client and a very different auth flow — ask if you
  need that variant too; this extension used to work that way before this
  rewrite).

## Why not just use Thunderbird's native calendar sync?

Thunderbird 156 ships the beginnings of native Graph/Exchange calendar
support (`GraphCalendar.sys.mjs` / `GraphProvider.sys.mjs`, backed by a
native `IExchangeClient` component tied to Thunderbird's own EWS
mail-account support). As of this build it's read-only (every write method
throws `NS_ERROR_NOT_IMPLEMENTED`), stores events in a transient in-memory
calendar that doesn't survive a restart, and isn't wired into any
user-facing UI yet (referenced bugs: 2052326, 2058691, 2058697). Worth
checking again in a future Thunderbird release — but for now it isn't
usable, which is why this extension exists.
