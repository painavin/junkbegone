const { PublicClientApplication } = require("@azure/msal-node");
const { BlobServiceClient } = require("@azure/storage-blob");

const CONTAINER_NAME = process.env.BLOB_CONTAINER || "junkbegone";
const TOKEN_CACHE_BLOB = "token-cache.json";
const SENDERS_BLOB = "junk-senders.json";

function getContainerClient() {
  const connectionString = process.env.AZURE_STORAGE_CONNECTION_STRING;
  if (!connectionString) {
    throw new Error("AZURE_STORAGE_CONNECTION_STRING is not set.");
  }
  const blobService = BlobServiceClient.fromConnectionString(connectionString);
  return blobService.getContainerClient(CONTAINER_NAME);
}

async function downloadBlobText(containerClient, blobName) {
  const blockBlob = containerClient.getBlockBlobClient(blobName);
  if (!(await blockBlob.exists())) return null;
  const buffer = await blockBlob.downloadToBuffer();
  return buffer.toString("utf-8");
}

async function uploadBlobText(containerClient, blobName, text) {
  const blockBlob = containerClient.getBlockBlobClient(blobName);
  await blockBlob.upload(text, Buffer.byteLength(text), { overwrite: true });
}

// MSAL cache plugin backed by a single blob, so refresh-token rotation
// is persisted automatically on every acquireTokenSilent call.
function createBlobCachePlugin(containerClient) {
  return {
    beforeCacheAccess: async (cacheContext) => {
      console.log("[cache] reading token-cache blob...");
      const cached = await downloadBlobText(containerClient, TOKEN_CACHE_BLOB);
      console.log("[cache] read complete. found existing cache:", !!cached);
      if (cached) cacheContext.tokenCache.deserialize(cached);
    },
    afterCacheAccess: async (cacheContext) => {
      console.log("[cache] afterCacheAccess. changed:", cacheContext.cacheHasChanged);
      if (cacheContext.cacheHasChanged) {
        console.log("[cache] writing token-cache blob...");
        await uploadBlobText(containerClient, TOKEN_CACHE_BLOB, cacheContext.tokenCache.serialize());
        console.log("[cache] write complete.");
      }
    },
  };
}

function createPca(containerClient) {
  return new PublicClientApplication({
    auth: {
      clientId: process.env.CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
    },
    cache: { cachePlugin: createBlobCachePlugin(containerClient) },
  });
}

async function getConservativeSenders(containerClient) {
  const text = await downloadBlobText(containerClient, SENDERS_BLOB);
  if (!text) return [];
  return JSON.parse(text);
}

async function getAccessToken(pca) {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length === 0) {
    throw new Error("No cached account. Run bootstrap-login.js once to sign in.");
  }
  const result = await pca.acquireTokenSilent({
    account: accounts[0],
    scopes: ["Mail.ReadWrite", "User.Read"],
  });
  return result.accessToken;
}

async function getJunkMessages(token) {
  const messages = [];
  let url =
    "https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages?$select=id,subject,from,flag&$top=100";

  while (url) {
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
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

function senderMatches(message, senders) {
  const from = message.from && message.from.emailAddress;
  if (!from) return false;
  const haystack = `${from.name || ""} ${from.address || ""}`;

  return senders.some((word) => {
    const regex = compileBadWord(word);
    return regex ? regex.test(haystack) : haystack.toLowerCase().includes(word.toLowerCase());
  });
}

// A follow-up flag on junk mail is treated as a manual "delete this" marker.
// Only an active flag counts; "complete" means the flag was already ticked off.
function isFlagged(message) {
  return !!message.flag && message.flag.flagStatus === "flagged";
}

// Words shorter than this match by accident: without it, "Bank of America"
// <alerts@bofa.com> counts as a match because "bofa" contains "of".
const NAME_WORD_MIN_LENGTH = 3;

// Display name split into comparable words: lowercased, accents stripped, any
// punctuation treated as a separator, and too-short tokens dropped.
function nameWords(name) {
  return name
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .split(/[^a-z0-9]+/)
    .filter((word) => word.length >= NAME_WORD_MIN_LENGTH);
}

// True when not one word of the display name appears anywhere in the address —
// "Costco Deals" <x7f2q@mailer-9.example>. Legitimate senders usually share at
// least one token between the two; spoofed ones typically share none.
//
// Deliberately false when there is no usable display name: no name means no words,
// and "no word appears" would then be vacuously true, matching every sender that
// omits a display name.
function nameMismatchesAddress(message) {
  const from = message.from && message.from.emailAddress;
  if (!from) return false;

  const name = (from.name || "").trim();
  const address = (from.address || "").toLowerCase();
  if (!name || !address) return false;
  if (name.toLowerCase() === address) return false;

  const words = nameWords(name);
  if (words.length === 0) return false;

  return !words.some((word) => address.includes(word));
}

function shouldDelete(message, senders) {
  return senderMatches(message, senders) || isFlagged(message) || nameMismatchesAddress(message);
}

// Mark read, then move to Deleted Items instead of deleting outright, so a false
// positive stays recoverable. The move has to come second: it returns the message
// under a new id, which would make a subsequent PATCH target a stale one.
//
// 404 is tolerated on both calls — the message already being gone is the desired
// end state.
async function moveToDeletedItems(token, id) {
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  const markRead = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}`, {
    method: "PATCH",
    headers,
    body: JSON.stringify({ isRead: true }),
  });
  if (!markRead.ok && markRead.status !== 404) {
    throw new Error(`Marking ${id} read failed: ${markRead.status} ${await markRead.text()}`);
  }

  const move = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${id}/move`, {
    method: "POST",
    headers,
    body: JSON.stringify({ destinationId: "deleteditems" }),
  });
  if (!move.ok && move.status !== 404) {
    throw new Error(`Moving ${id} to Deleted Items failed: ${move.status} ${await move.text()}`);
  }
}

async function runCleanup() {
  const containerClient = getContainerClient();
  const pca = createPca(containerClient);

  const token = await getAccessToken(pca);
  const senders = await getConservativeSenders(containerClient);

  const messages = await getJunkMessages(token);
  const toMove = messages.filter((m) => shouldDelete(m, senders));
  const flagged = toMove.filter(isFlagged).length;
  const mismatched = toMove.filter(nameMismatchesAddress).length;

  let movedCount = 0;
  for (const message of toMove) {
    await moveToDeletedItems(token, message.id);
    movedCount++;
  }

  return {
    scanned: messages.length,
    matched: toMove.length,
    flagged,
    mismatched,
    moved: movedCount,
  };
}

module.exports = { createPca, getContainerClient, getConservativeSenders, runCleanup, CONTAINER_NAME, TOKEN_CACHE_BLOB, SENDERS_BLOB };
