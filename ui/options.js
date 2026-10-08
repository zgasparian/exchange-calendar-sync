const urlInput = document.getElementById("url");
const usernameInput = document.getElementById("username");
const passwordInput = document.getElementById("password");
const connectButton = document.getElementById("connect");
const loginStatus = document.getElementById("loginStatus");
const accountsList = document.getElementById("accounts");
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

async function refreshAccounts() {
  const accounts = await browser.runtime.sendMessage({ type: "listAccounts" });
  accountsList.innerHTML = "";
  if (!accounts.length) {
    const li = document.createElement("li");
    li.textContent = "No accounts connected yet.";
    accountsList.appendChild(li);
    return;
  }
  for (const account of accounts) {
    const li = document.createElement("li");

    const meta = document.createElement("div");
    meta.className = "acct-meta";
    const name = document.createElement("span");
    name.textContent = account.username;
    meta.appendChild(name);
    const url = document.createElement("span");
    url.className = "acct-url";
    url.textContent = account.url;
    meta.appendChild(url);
    const synced = document.createElement("span");
    synced.className = "acct-url";
    synced.textContent = formatLastSynced(account.lastSyncedAt);
    meta.appendChild(synced);
    if (account.needsReauth) {
      const warn = document.createElement("span");
      warn.className = "warn";
      warn.textContent = "Sign-in failed — reconnect below with your current password.";
      meta.appendChild(warn);
    }
    li.appendChild(meta);

    const removeButton = document.createElement("button");
    removeButton.className = "remove";
    removeButton.textContent = "Remove";
    removeButton.addEventListener("click", async () => {
      removeButton.disabled = true;
      await browser.runtime.sendMessage({ type: "removeAccount", accountId: account.accountId });
      await refreshAccounts();
    });
    li.appendChild(removeButton);

    accountsList.appendChild(li);
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
