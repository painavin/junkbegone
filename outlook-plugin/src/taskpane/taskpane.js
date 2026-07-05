/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, window, Office, fetch */

import { PublicClientApplication } from "@azure/msal-browser";

const DEFAULT_BAD_WORDS = ["AARP", "AccuQuote", "Affordable"];
const ROAMING_SETTINGS_KEY = "junker.badWords";

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
    document.getElementById("badwords").value = loadBadWords().join("\n");

    applyOfficeTheme();
    if (Office.context.officeTheme) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyOfficeTheme);
    }
  }
});

function applyOfficeTheme() {
  const theme = Office.context.officeTheme;
  if (!theme) return;

  const root = document.documentElement.style;
  root.setProperty("--bg", theme.bodyBackgroundColor);
  root.setProperty("--fg", theme.bodyForegroundColor);
}

function loadBadWords() {
  const stored = Office.context.roamingSettings.get(ROAMING_SETTINGS_KEY);
  return stored && stored.length ? stored : DEFAULT_BAD_WORDS;
}

function saveBadWordsFromTextarea() {
  const words = document
    .getElementById("badwords")
    .value.split("\n")
    .map((w) => w.trim())
    .filter((w) => w.length > 0);

  Office.context.roamingSettings.set(ROAMING_SETTINGS_KEY, words);
  Office.context.roamingSettings.saveAsync((result) => {
    const resultEl = document.getElementById("result");
    resultEl.textContent =
      result.status === Office.AsyncResultStatus.Succeeded
        ? `Saved ${words.length} bad word(s).`
        : `Failed to save: ${result.error.message}`;
  });
}

async function getAccessToken() {
  await msalInstance.initialize();
  const scopes = ["Mail.ReadWrite", "User.Read"];

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await msalInstance.acquireTokenSilent({ scopes, account: accounts[0] });
      return result.accessToken;
    } catch {
      // fall through to interactive login
    }
  }

  const result = await msalInstance.loginPopup({ scopes });
  return result.accessToken;
}

async function getJunkMessages(token) {
  const messages = [];
  let url =
    "https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages?$select=id,subject,from&$top=100";

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
    const badWords = loadBadWords();

    resultEl.textContent = "Scanning Junk Email folder...";
    const messages = await getJunkMessages(token);
    const matched = messages.filter((m) => senderMatchesBadWord(m, badWords));

    const lines = matched.map((m) => `  ${m.from.emailAddress.name} <${m.from.emailAddress.address}> — ${m.subject}`);
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
    const badWords = loadBadWords();

    resultEl.textContent = "Scanning Junk Email folder...";
    const messages = await getJunkMessages(token);
    const toDelete = messages.filter((m) => senderMatchesBadWord(m, badWords));

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
