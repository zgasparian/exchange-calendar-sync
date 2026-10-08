const urlInput = document.getElementById("url");
const usernameInput = document.getElementById("username");
const passwordInput = document.getElementById("password");
const connectButton = document.getElementById("connect");
const loginStatus = document.getElementById("loginStatus");
const accountsContainer = document.getElementById("accounts");
const syncIntervalInput = document.getElementById("syncInterval");
const saveIntervalButton = document.getElementById("saveInterval");
const intervalStatus = document.getElementById("intervalStatus");

function setStatus(el, message, kind) {
  el.textContent = message;
  el.className = `status ${kind || ""}`;
}

function formatLastSynced(lastSyncedAt) {
  return lastSyncedAt ? `Last synced: ${new Date(lastSyncedAt).toLocaleString()}` : "Last synced: never";
}

function renderCalendarRow(calendar) {
  const row = document.createElement("div");
  row.className = "calendar-row";

  const meta = document.createElement("div");
  meta.className = "calendar-meta";
  const name = document.createElement("span");
  name.textContent = calendar.folderName;
  meta.appendChild(name);
  const synced = document.createElement("span");
  synced.className = "synced-at";
  synced.textContent = formatLastSynced(calendar.lastSyncedAt);
  meta.appendChild(synced);
  if (calendar.needsReauth) {
    const warn = document.createElement("span");
    warn.className = "warn";
    warn.textContent = "Sign-in failed — fix the account's password below to resume syncing.";
    meta.appendChild(warn);
  }
  row.appendChild(meta);

  const removeButton = document.createElement("button");
  removeButton.className = "remove small";
  removeButton.textContent = "Remove";
  removeButton.addEventListener("click", async () => {
    removeButton.disabled = true;
    await browser.runtime.sendMessage({ type: "removeCalendar", calendarKey: calendar.calendarKey });
    await refreshAccounts();
  });
  row.appendChild(removeButton);

  return row;
}

function renderAccount(account) {
  const block = document.createElement("div");
  block.className = "account-block";

  const header = document.createElement("div");
  header.className = "account-header";

  const meta = document.createElement("div");
  meta.className = "acct-meta";
  const name = document.createElement("span");
  name.textContent = account.username;
  meta.appendChild(name);
  const url = document.createElement("span");
  url.className = "acct-url";
  url.textContent = account.url;
  meta.appendChild(url);
  if (account.needsReauth) {
    const warn = document.createElement("span");
    warn.className = "warn";
    warn.textContent = "Sign-in failed — reconnect above with this account's current password.";
    meta.appendChild(warn);
  }
  header.appendChild(meta);

  const actions = document.createElement("div");
  const addCalendarButton = document.createElement("button");
  addCalendarButton.className = "small";
  addCalendarButton.textContent = "Add calendar…";
  actions.appendChild(addCalendarButton);
  const removeAccountButton = document.createElement("button");
  removeAccountButton.className = "remove small";
  removeAccountButton.textContent = "Remove account";
  actions.appendChild(removeAccountButton);
  header.appendChild(actions);

  block.appendChild(header);

  for (const calendar of account.calendars) {
    block.appendChild(renderCalendarRow(calendar));
  }

  const discoveryArea = document.createElement("div");
  discoveryArea.className = "add-calendar-area";
  block.appendChild(discoveryArea);

  addCalendarButton.addEventListener("click", async () => {
    addCalendarButton.disabled = true;
    discoveryArea.textContent = "Looking up calendars on this account…";
    const response = await browser.runtime.sendMessage({ type: "listFolders", accountId: account.accountId });
    addCalendarButton.disabled = false;
    discoveryArea.textContent = "";
    if (!response.ok) {
      discoveryArea.textContent = `Could not list calendars: ${response.error}`;
      return;
    }
    const unsynced = response.folders.filter(f => !f.synced);
    if (!unsynced.length) {
      discoveryArea.textContent = "Every calendar on this account is already synced.";
      return;
    }
    for (const folder of unsynced) {
      const row = document.createElement("div");
      row.className = "folder-row";
      const label = document.createElement("span");
      label.textContent = folder.name;
      row.appendChild(label);
      const addButton = document.createElement("button");
      addButton.className = "small";
      addButton.textContent = "Add";
      addButton.addEventListener("click", async () => {
        addButton.disabled = true;
        await browser.runtime.sendMessage({
          type: "addCalendar",
          accountId: account.accountId,
          folderRef: folder.folderRef,
          folderName: folder.name,
        });
        await refreshAccounts();
      });
      row.appendChild(addButton);
      discoveryArea.appendChild(row);
    }
  });

  removeAccountButton.addEventListener("click", async () => {
    removeAccountButton.disabled = true;
    await browser.runtime.sendMessage({ type: "removeAccount", accountId: account.accountId });
    await refreshAccounts();
  });

  return block;
}

async function refreshAccounts() {
  const accounts = await browser.runtime.sendMessage({ type: "listAccounts" });
  accountsContainer.innerHTML = "";
  if (!accounts.length) {
    const empty = document.createElement("p");
    empty.className = "hint";
    empty.textContent = "No accounts connected yet.";
    accountsContainer.appendChild(empty);
    return;
  }
  for (const account of accounts) {
    accountsContainer.appendChild(renderAccount(account));
  }
}

connectButton.addEventListener("click", async () => {
  const url = urlInput.value.trim();
  const username = usernameInput.value.trim();
  const password = passwordInput.value;
  if (!url || !username || !password) {
    setStatus(loginStatus, "Fill in the server URL, username, and password.", "error");
    return;
  }
  connectButton.disabled = true;
  setStatus(loginStatus, "Connecting…", "");
  const response = await browser.runtime.sendMessage({ type: "connectAccount", url, username, password });
  connectButton.disabled = false;
  if (response.ok) {
    setStatus(loginStatus, `Connected ${response.account.username}.`, "ok");
    passwordInput.value = "";
    await refreshAccounts();
  } else {
    setStatus(loginStatus, `Connection failed: ${response.error}`, "error");
  }
});

async function loadSyncInterval() {
  const { minutes } = await browser.runtime.sendMessage({ type: "getSyncIntervalMinutes" });
  syncIntervalInput.value = minutes;
}

saveIntervalButton.addEventListener("click", async () => {
  const minutes = parseInt(syncIntervalInput.value, 10);
  if (!minutes || minutes < 1) {
    setStatus(intervalStatus, "Enter a whole number of minutes (1 or more).", "error");
    return;
  }
  const response = await browser.runtime.sendMessage({ type: "setSyncIntervalMinutes", minutes });
  syncIntervalInput.value = response.minutes;
  setStatus(intervalStatus, `Saved — syncing every ${response.minutes} minute${response.minutes === 1 ? "" : "s"}.`, "ok");
});

refreshAccounts();
loadSyncInterval();
