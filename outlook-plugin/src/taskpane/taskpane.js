/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, window, Office, fetch */

import { PublicClientApplication } from "@azure/msal-browser";

// Only used when the blob does not exist yet.
const DEFAULT_BAD_WORDS = ["AARP", "AccuQuote", "Affordable"];

// The bad-word list lives in the blob the timer-triggered function reads, so the
// add-in and the scheduled cleanup always act on the same list.
const STORAGE_ACCOUNT = "satickers";
const BLOB_CONTAINER = "junkbegone";
const SENDERS_BLOB = "junk-senders.json";
const SENDERS_BLOB_URL = `https://${STORAGE_ACCOUNT}.blob.core.windows.net/${BLOB_CONTAINER}/${SENDERS_BLOB}`;
const BLOB_API_VERSION = "2021-08-06";

// AAD issues one token per resource, so Graph and Storage need separate requests.
const GRAPH_SCOPES = ["Mail.ReadWrite", "User.Read"];
const STORAGE_SCOPES = ["https://storage.azure.com/user_impersonation"];

// ETag of the list as we last read it, used to detect a competing write on save.
let sendersETag = null;
// Whether the textarea reflects the stored list. Saving before it does would
// overwrite the real list with whatever placeholder text is on screen.
let badWordsLoaded = false;

const msalConfig = {
  auth: {
    clientId: "ed7fafe8-a6fc-4ca8-ae75-db77f33f2c5f",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + window.location.pathname,
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("preview").onclick = preview;
    document.getElementById("save-badwords").onclick = saveBadWordsFromTextarea;
    populateBadWords();

    applyOfficeTheme();
    if (Office.context.officeTheme) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyOfficeTheme);
    }
  }
});

// Runs without a user gesture, so a sign-in popup here would be blocked — only a
// cached token can be used. Preview, Run Cleanup and Save List sign in interactively.
async function populateBadWords() {
  const resultEl = document.getElementById("result");
  try {
    const words = await loadBadWords({ interactive: false });
    if (!words) {
      resultEl.textContent = "Sign in with Preview or Run Cleanup to load the saved list.";
      return;
    }
    showBadWords(words);
  } catch (error) {
    resultEl.textContent = `Could not load the saved list: ${error.message}`;
  }
}

function applyOfficeTheme() {
  const theme = Office.context.officeTheme;
  if (!theme) return;

  const root = document.documentElement.style;
  root.setProperty("--bg", theme.bodyBackgroundColor);
  root.setProperty("--fg", theme.bodyForegroundColor);
}

function showBadWords(words) {
  document.getElementById("badwords").value = words.join("\n");
  badWordsLoaded = true;
}

// Returns the stored list, or null when interactive sign-in was declined/skipped.
async function loadBadWords({ interactive = true } = {}) {
  const token = await acquireToken(STORAGE_SCOPES, { interactive });
  if (!token) return null;

  const response = await fetch(SENDERS_BLOB_URL, {
    headers: { Authorization: `Bearer ${token}`, "x-ms-version": BLOB_API_VERSION },
  });

  if (response.status === 404) {
    sendersETag = null;
    return DEFAULT_BAD_WORDS;
  }
  if (!response.ok) {
    throw new Error(`Reading ${SENDERS_BLOB} failed: ${response.status} ${await response.text()}`);
  }

  sendersETag = response.headers.get("ETag");
  const words = JSON.parse(await response.text());
  return Array.isArray(words) ? words : DEFAULT_BAD_WORDS;
}

async function saveBadWordsFromTextarea() {
  const resultEl = document.getElementById("result");

  if (!badWordsLoaded) {
    resultEl.textContent =
      "The saved list has not loaded yet — saving now would overwrite it. Run Preview first, then save.";
    return;
  }

  const words = document
    .getElementById("badwords")
    .value.split("\n")
    .map((w) => w.trim())
    .filter((w) => w.length > 0);

  resultEl.textContent = "Saving list...";

  try {
    const token = await acquireToken(STORAGE_SCOPES);
    const headers = {
      Authorization: `Bearer ${token}`,
      "x-ms-version": BLOB_API_VERSION,
      "x-ms-blob-type": "BlockBlob",
      "Content-Type": "application/json",
    };
    // Fail rather than silently clobber a list edited elsewhere since we read it.
    if (sendersETag) headers["If-Match"] = sendersETag;

    const response = await fetch(SENDERS_BLOB_URL, {
      method: "PUT",
      headers,
      body: JSON.stringify(words, null, 2),
    });

    if (response.status === 412) {
      resultEl.textContent =
        "The saved list changed elsewhere. Reopen the task pane to reload it, then re-apply your edits.";
      return;
    }
    if (!response.ok) {
      throw new Error(`${response.status} ${await response.text()}`);
    }

    sendersETag = response.headers.get("ETag");
    resultEl.textContent = `Saved ${words.length} bad word(s). The scheduled cleanup will use this list too.`;
  } catch (error) {
    resultEl.textContent = `Failed to save: ${error.message}`;
  }
}

async function acquireToken(scopes, { interactive = true } = {}) {
  await msalInstance.initialize();

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await msalInstance.acquireTokenSilent({ scopes, account: accounts[0] });
      return result.accessToken;
    } catch {
      if (!interactive) return null;
      const result = await msalInstance.acquireTokenPopup({ scopes, account: accounts[0] });
      return result.accessToken;
    }
  }

  if (!interactive) return null;
  const result = await msalInstance.loginPopup({ scopes });
  return result.accessToken;
}

function getAccessToken() {
  return acquireToken(GRAPH_SCOPES);
}

async function getJunkMessages(token) {
  const messages = [];
  let url =
    "https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages?$select=id,subject,from,flag&$top=100";

  while (url) {
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) {
      throw new Error(`Graph request failed: ${response.status} ${await response.text()}`);
    }
    const body = await response.json();
    messages.push(...body.value);
    url = body["@odata.nextLink"] || null;
  }

  return messages;
}

const REGEX_LITERAL = /^\/(.*)\/([a-z]*)$/;

function compileBadWord(word) {
  const literal = word.match(REGEX_LITERAL);
  if (!literal) return null;
  try {
    return new RegExp(literal[1], literal[2]);
  } catch {
    return null;
  }
}

function senderMatchesBadWord(message, badWords) {
  const from = message.from && message.from.emailAddress;
  if (!from) return false;
  const haystack = `${from.name || ""} ${from.address || ""}`;

  return badWords.some((word) => {
    const regex = compileBadWord(word);
    return regex ? regex.test(haystack) : haystack.toLowerCase().includes(word.toLowerCase());
  });
}

// A follow-up flag on junk mail is treated as a manual "delete this" marker.
// Only an active flag counts; "complete" means the flag was already ticked off.
function isFlagged(message) {
  return !!message.flag && message.flag.flagStatus === "flagged";
}

function shouldDelete(message, badWords) {
  return senderMatchesBadWord(message, badWords) || isFlagged(message);
}

function matchReason(message, badWords) {
  const reasons = [];
  if (senderMatchesBadWord(message, badWords)) reasons.push("bad word");
  if (isFlagged(message)) reasons.push("flagged");
  return reasons.join(" + ");
}

// Flagged messages match regardless of sender, so `from` may be absent here.
function describeSender(message) {
  const from = message.from && message.from.emailAddress;
  if (!from) return "(no sender)";
  return `${from.name || ""} <${from.address || ""}>`;
}

async function deleteMessage(token, id) {
  const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok && response.status !== 404) {
    throw new Error(`Delete failed for ${id}: ${response.status} ${await response.text()}`);
  }
}

export async function preview() {
  const resultEl = document.getElementById("result");
  resultEl.textContent = "Signing in...";

  try {
    const token = await getAccessToken();
    const badWords = await loadBadWords();
    if (!badWordsLoaded) showBadWords(badWords);

    resultEl.textContent = "Scanning Junk Email folder...";
    const messages = await getJunkMessages(token);
    const matched = messages.filter((m) => shouldDelete(m, badWords));

    const lines = matched.map((m) => `  [${matchReason(m, badWords)}] ${describeSender(m)} — ${m.subject}`);
    resultEl.textContent =
      `Scanned: ${messages.length} messages. Would match: ${matched.length}.\n` + lines.join("\n");
  } catch (error) {
    resultEl.textContent = `Error: ${error.message}`;
  }
}

export async function run() {
  const resultEl = document.getElementById("result");
  resultEl.textContent = "Signing in...";

  try {
    const token = await getAccessToken();
    const badWords = await loadBadWords();
    if (!badWordsLoaded) showBadWords(badWords);

    resultEl.textContent = "Scanning Junk Email folder...";
    const messages = await getJunkMessages(token);
    const toDelete = messages.filter((m) => shouldDelete(m, badWords));

    let deleteCount = 0;
    for (const message of toDelete) {
      resultEl.textContent = `Deleting ${deleteCount + 1} of ${toDelete.length}...`;
      await deleteMessage(token, message.id);
      deleteCount++;
    }

    resultEl.textContent = `Scanned: ${messages.length} messages. Matched: ${toDelete.length}. Deleted: ${deleteCount}.`;
  } catch (error) {
    resultEl.textContent = `Error: ${error.message}`;
  }
}
