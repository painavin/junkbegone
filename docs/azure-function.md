# Azure Function

A timer-triggered Azure Function that deletes unwanted mail from the **Junk Email** folder on a
schedule, with no user interaction. It's the unattended counterpart to the
[Outlook plugin](./outlook-plugin.md).

Source lives in [`az-function/`](../az-function).

## Deployed resources

| Thing | Value |
| --- | --- |
| Function app | `junkbegone` (Flex Consumption) |
| Resource group | `rg-sub-free` |
| Storage account | `satickers` |
| Blob container | `junkbegone` |
| App registration | `ed7fafe8-a6fc-4ca8-ae75-db77f33f2c5f` (shared with the add-in) |

The subscription is deliberately not recorded here — look it up with `az account show --query id -o tsv`.

Flex Consumption apps get a random suffix in their hostname, so look the hostname up with the CLI
rather than hardcoding it (see [Triggering manually](#triggering-manually)).

## How it runs

`dailyCleanup` is registered with the Node.js v4 programming model in
[`src/functions/dailyCleanup.js`](../az-function/src/functions/dailyCleanup.js):

```js
app.timer("dailyCleanup", { schedule: "0 0 */12 * * *", ... })
```

That NCRONTAB expression is **every 12 hours** (six fields — seconds first — so `0 0 */12 * * *` is
"second 0, minute 0, every 12th hour"). Despite the function's name it is not daily. The handler logs
`Scanned / Matched (flagged) / Deleted` counts and rethrows on failure so the invocation is recorded
as failed.

All real work is in [`graphClient.js`](../az-function/graphClient.js) via `runCleanup()`, which
returns `{ scanned, matched, flagged, deleted }`.

## Deletion rules

Identical to the add-in's — a junk message is deleted if **either** matches:

1. **Sender matches an entry in `junk-senders.json`**, tested against
   `"<sender display name> <sender address>"`. Plain entries are case-insensitive substring matches;
   `/pattern/flags` entries are compiled as regexes, falling back to substring matching if malformed.
2. **The message carries an active follow-up flag** (`flag.flagStatus === "flagged"`). `complete` and
   `notFlagged` are ignored.

See `senderMatches`, `isFlagged`, and `shouldDelete` in `graphClient.js`. Keep these in sync with the
add-in's copies in [`taskpane.js`](../outlook-plugin/src/taskpane/taskpane.js) — the *list* is shared
via the blob, but the matching *logic* is duplicated in both codebases, not shared.

> **Careful:** flagging junk mail marks it for deletion rather than preservation. Move a false
> positive out of the Junk folder instead of flagging it.

Messages are read from `/me/mailFolders/junkemail/messages` with
`$select=id,subject,from,flag&$top=100`, paging through `@odata.nextLink` until the folder is
exhausted. `flag` must stay in `$select` or every message looks unflagged. Deletion is
`DELETE /me/messages/{id}`, treating `404` as success.

## Authentication

The function runs as *you*, using delegated permissions — there is no app-only Graph access here.
That requires a refresh token to survive between invocations, which works as follows:

1. **One-time bootstrap.** `node bootstrap-login.js` (`npm run bootstrap-login`) signs you in with
   the MSAL device-code flow, creates the container if absent, seeds a placeholder
   `junk-senders.json` if absent, and writes the serialised MSAL token cache to blob
   `token-cache.json`.
2. **Every run after that.** `createPca()` builds an MSAL `PublicClientApplication` with a custom
   cache plugin (`createBlobCachePlugin`) whose `beforeCacheAccess` reads `token-cache.json` and
   whose `afterCacheAccess` writes it back whenever the cache changed. So refresh-token rotation is
   persisted automatically and `acquireTokenSilent` keeps working indefinitely.

If the refresh token ever expires or is revoked, `getAccessToken` throws
`No cached account. Run bootstrap-login.js once to sign in.` — re-run the bootstrap to recover.

## Blobs

Both live in container `junkbegone` on `satickers`:

| Blob | Purpose |
| --- | --- |
| `junk-senders.json` | JSON array of strings — the bad-sender patterns. Read by the function, read *and written* by the add-in's **Save List**. |
| `token-cache.json` | Serialised MSAL token cache. **Contains a refresh token — treat as a secret.** |

This blob is the single source of truth for the list: the [Outlook plugin](./outlook-plugin.md) edits
it directly from the browser, so the add-in and this function always act on the same patterns.

`junk-senders.json` is a flat array. Blank lines are used to group entries alphabetically for human
readability, and survive round-tripping through the CLI — but note the add-in rewrites the file with
`JSON.stringify(words, null, 2)`, which strips them:

```json
[
    "/AARP.*/",
    "Americor",

    "Brinks",
    "BetterHelp"
]
```

A copy is kept at [`az-function/junk-senders.json`](../az-function/junk-senders.json) for editing and
diffing in git, but **the function only ever reads the blob** — editing the local file has no effect
until you upload it.

### Download

```
az storage blob download --account-name satickers --container-name junkbegone \
  --name junk-senders.json --file ./junk-senders.json --auth-mode login
```

### Upload

```
az storage blob upload --account-name satickers --container-name junkbegone \
  --name junk-senders.json --file ./junk-senders.json --overwrite --auth-mode login
```

`--auth-mode login` needs a **Storage Blob Data Contributor** role assignment on the container (data
plane); being subscription Owner is not sufficient. Drop `--auth-mode login` to fall back to the
account key, which the CLI will fetch automatically if you have control-plane rights. Newly created
role assignments can take a few minutes to take effect on the data plane.

## Configuration

Locally, `local.settings.json` (gitignored — see `local.settings.json.example`); in Azure, the
function app's application settings. `bootstrap-login.js` additionally reads `.env` via `dotenv`.

| Setting | Purpose |
| --- | --- |
| `AZURE_STORAGE_CONNECTION_STRING` | Storage account used for the two blobs above. |
| `BLOB_CONTAINER` | Container name; defaults to `junkbegone`. |
| `CLIENT_ID` | App registration used for the delegated Graph token. |
| `AzureWebJobsStorage` | Functions runtime's own storage. |
| `FUNCTIONS_WORKER_RUNTIME` | `node`. |

`local.settings.json` and `.env` hold a storage account key. Both are gitignored — keep it that way,
and never paste the key into docs, commit messages, or issues.

## Local development

```
cd az-function
npm install
npm run bootstrap-login   # one time, seeds the blob token cache
npm start                 # func start
```

`npm start` requires the Azure Functions Core Tools (`func`). A timer trigger won't fire immediately
on start; use the admin endpoint to invoke it on demand.

## Deployment

```
cd az-function
func azure functionapp publish junkbegone
```

## Triggering manually

```
HOST=$(az functionapp show --name junkbegone --resource-group rg-sub-free --query defaultHostName -o tsv)
MASTER_KEY=$(az functionapp keys list --name junkbegone --resource-group rg-sub-free --query masterKey -o tsv)
curl -X POST "https://$HOST/admin/functions/dailyCleanup?code=$MASTER_KEY" \
  -H "Content-Type: application/json" -d '{"input": ""}'
```

This deletes mail immediately — there is no preview mode on the function side. To see what would
match without deleting anything, use **Preview** in the add-in.

## Logs

```
az webapp log tail --name junkbegone --resource-group rg-sub-free
```

Or the **Log stream** blade in the portal. Application Insights is enabled with sampling, excluding
`Request` telemetry (see `host.json`).
