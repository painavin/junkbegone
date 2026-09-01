# Outlook Plugin

An Office.js mail add-in that scans the **Junk Email** folder and moves unwanted messages to Deleted
Items, on demand (Preview / Run Cleanup) from a task pane.

Source lives in [`outlook-plugin/`](../outlook-plugin).

## What it does

The add-in shows a task pane with an editable bad-word list and these buttons:

| Button | Behaviour |
| --- | --- |
| **Sign in** | Signs in and loads the list. Only shown when there's no cached account. |
| **Save List** | Writes the list back to the shared blob. |
| **Preview** | Scans Junk Email and lists what *would* be moved, with the reason for each match. Changes nothing. |
| **Run Cleanup** | Scans Junk Email and moves every matching message to Deleted Items, marking it read. |

Use **Preview** first after editing the list. Matches are moved to Deleted Items rather than purged, so
a false positive is recoverable from there until that folder is emptied.

## Deletion rules

A junk message is moved to Deleted Items if **either** rule matches:

1. **Sender matches the bad-word list** — each entry is tested against the string
   `"<sender display name> <sender address>"`.
   - A plain entry (`Americor`) is a case-insensitive substring match.
   - An entry wrapped in slashes (`/Credit.*Card/`, `/Cloud.*Stor/i`) is compiled as a JavaScript
     regex, with optional trailing flags. A malformed regex falls back to substring matching rather
     than throwing.
2. **The message carries an active follow-up flag** — i.e. Graph reports
   `flag.flagStatus === "flagged"`.

Rule 2 lets you flag junk by hand in Outlook and have the next run move it, without adding a
pattern to the list. Only an *active* flag counts; `complete` (a flag you've already ticked off) and
`notFlagged` are ignored.

> **Careful:** a follow-up flag conventionally means "keep this, follow up on it". Here it means the
> opposite. Do not use flags to rescue a false positive sitting in Junk — flagging it marks it for
> removal. Move it out of the Junk folder instead.

Because a flagged message matches regardless of who sent it, matches may have no sender at all; the
preview renders those as `(no sender)`.

Relevant code in [`src/taskpane/taskpane.js`](../outlook-plugin/src/taskpane/taskpane.js):
`senderMatchesBadWord`, `isFlagged`, `shouldDelete`, `matchReason`.

## The bad-word list

The list is stored as `junk-senders.json` in Blob Storage — the same blob the Azure Function reads.
It is the single source of truth: editing it in the task pane changes what the scheduled cleanup does,
and **Preview** shows exactly what the scheduled cleanup would move. See
[azure-function.md](./azure-function.md) for the blob's location and format.

The add-in reads the blob over the Blob REST API with plain `fetch` (no `@azure/storage-blob`, to keep
the bundle small) and writes it back on **Save List**. `DEFAULT_BAD_WORDS` is only a fallback for when
the blob does not exist yet.

A local copy is kept at [`az-function/junk-senders.json`](../az-function/junk-senders.json) for
git-diffable editing, but it is only a copy — the blob is what both components read.

### Loading and the sign-in constraint

`Office.onReady` fires without a user gesture, so a sign-in popup there would be blocked by the
browser. `populateBadWords()` therefore requests a **silent-only** token, which gives two paths:

- **Cached account** — the list loads immediately, the pane shows *"Signed in as …"*, and the
  **Sign in** button stays hidden.
- **No cached account** — the **Sign in** button appears. It runs from a click, so it may sign in
  interactively. It acquires the Graph *and* Storage tokens together and then loads the list.

Acquiring both tokens in `signIn()` is deliberate: Graph and Storage use different authorities and
consent separately, so the first sign-in can produce two popups back to back. Doing it in one place
keeps that burst at the moment you asked to sign in, instead of a second popup appearing later when
you click Preview. Preview and Run Cleanup can still sign in on their own if you skip the button.

The signed-in account is displayed because two identities are in play here — the mailbox is a personal
Microsoft account, while the blob is reached through that account's guest identity in another tenant —
so "which account am I?" is a real question when something fails.

### Guards against losing the list

Two failure modes are handled explicitly, because both silently destroy the list:

- **Saving before loading.** If the blob hasn't loaded, the textarea holds placeholder text rather
  than the real list, so saving would overwrite it. `saveBadWordsFromTextarea()` refuses, and says to
  sign in first. `badWordsLoaded` tracks this.
- **Concurrent edits.** The `ETag` from the read is sent back as `If-Match` on the write, so a list
  changed elsewhere since it was read produces a `412` instead of a silent clobber. The pane then
  asks you to reload and re-apply. `sendersETag` tracks this.

Reading the `ETag` response header from JavaScript requires the storage account's CORS rule to expose
it — the rule below uses `--exposed-headers "*"`, which covers it.

## Authentication

MSAL Browser (`@azure/msal-browser`) against the `common` authority, using app registration
`ed7fafe8-a6fc-4ca8-ae75-db77f33f2c5f`.

Two resources are involved, and AAD issues one access token per resource, so `acquireToken(scopes)`
is called twice with different scopes:

| Resource | Scopes | Used for |
| --- | --- | --- |
| Microsoft Graph | `Mail.ReadWrite`, `User.Read` | reading, marking read, and moving junk mail |
| Azure Storage | `https://storage.azure.com/user_impersonation` | reading and writing the list blob |

The two use **different authorities**, which matters:

- Graph uses `/common`, so the personal Microsoft account mailbox can sign in.
- Storage uses the tenant explicitly (`https://login.microsoftonline.com/<tenant-id>`), because Azure
  Storage rejects personal accounts. Via `/common` an MSA resolves to its consumer home tenant, which
  owns no Azure resources, and sign-in fails with *"You can't sign in here with a personal account.
  Use your work or school account instead."* The mailbox exists as a guest in the tenant owning the
  storage account, and that guest holds the RBAC role, so the token must be requested from that
  tenant.

Because the two authorities produce two MSAL account entries in different tenants, `pickAccount()`
selects by `tenantId` rather than taking `getAllAccounts()[0]`.

`acquireToken` tries `acquireTokenSilent` for a cached account, then falls back to
`acquireTokenPopup` (or `loginPopup` when there is no account yet) — unless called with
`{ interactive: false }`, in which case it returns `null` rather than opening a popup.

The redirect URI is computed at runtime as `window.location.origin + window.location.pathname`, so
both origins must be registered as SPA redirect URIs on the app registration:

- `https://localhost:3000/taskpane.html`
- `https://painavin.github.io/junkbegone/taskpane.html`

### Azure prerequisites for blob access

The add-in talks to Blob Storage directly from the browser, authenticated as the signed-in user —
there is no backend and no embedded SAS token or account key. That requires all four of:

1. **Delegated permission** — `user_impersonation` on Azure Storage
   (`e406a681-f3d4-42a8-90b6-c2b029497af1`), granted on the app registration and consented.
2. **RBAC** — the signing-in user holds **Storage Blob Data Contributor** on the `junkbegone`
   container. Subscription Owner is *not* sufficient; this is a data-plane role.

   ```
   az role assignment create --assignee-object-id <your-object-id> --assignee-principal-type User \
     --role "Storage Blob Data Contributor" \
     --scope "/subscriptions/<sub>/resourceGroups/rg-sub-free/providers/Microsoft.Storage/storageAccounts/satickers/blobServices/default/containers/junkbegone"
   ```

   Get the object id with `az ad signed-in-user show --query id -o tsv`. Pass it via
   `--assignee-object-id`, not `--assignee <email>` — for a personal Microsoft account signed in as a
   guest, the directory UPN is `navin.pai_outlook.com#EXT#@navinpai.onmicrosoft.com`, so lookup by
   email address fails.

   **Get the scope exactly right.** Container-scoped assignments are *not* validated against the
   parent resource group, so a scope naming a nonexistent resource group is accepted and returns
   clean JSON while granting nothing. The only symptom is `403
   AuthorizationPermissionMismatch` on every data-plane call, which looks identical to a propagation
   delay or an unsupported-identity problem. Verify with:

   ```
   az storage blob download --account-name satickers --container-name junkbegone \
     --name junk-senders.json --file ./rbac-test.json --auth-mode login
   ```

   That exercises the same principal and role as the add-in, so it isolates RBAC from anything
   browser-, CORS-, or MSAL-related.

   For the record: a personal Microsoft account signed in as a B2B guest **does** work for Storage
   data-plane RBAC. The guest object in the resource tenant is a normal Entra principal and can hold
   the role.
3. **CORS** on the storage account's Blob service, allowing both add-in origins:

   ```
   az storage cors add --services b --methods GET PUT OPTIONS HEAD \
     --origins "https://localhost:3000" "https://painavin.github.io" \
     --allowed-headers "*" --exposed-headers "*" --max-age 3600 \
     --account-name satickers
   ```

4. **The `x-ms-version` header** on every request (`2021-08-06` here), plus `x-ms-blob-type:
   BlockBlob` on writes. The REST API rejects requests without them.

Without the CORS rule the requests fail in the browser with an opaque CORS error rather than a
useful status code — check that first if the list won't load.

## Layout

```
outlook-plugin/
├── manifest.xml              add-in manifest (URLs point at localhost for dev)
├── webpack.config.js         build + dev server; rewrites dev URLs to prod on production builds
└── src/
    ├── taskpane/
    │   ├── taskpane.html     task pane markup
    │   ├── taskpane.css      task pane styles
    │   └── taskpane.js       all add-in logic: auth, Graph calls, matching, deletion
    └── commands/
        ├── commands.html     host page for the ribbon function file
        └── commands.js       ribbon "Perform an action" button (placeholder notification)
```

The task pane follows the host's light/dark theme via `Office.context.officeTheme`, mapped onto the
`--bg` / `--fg` CSS custom properties and refreshed on `OfficeThemeChanged`.

## Graph usage

Messages are read from `/me/mailFolders/junkemail/messages` with
`$select=id,subject,from,flag&$top=100`, following `@odata.nextLink` until exhausted — so the whole
folder is scanned, not just the first page. `flag` must stay in `$select`; Graph omits it otherwise
and every message would look unflagged.

Deletion is `DELETE /me/messages/{id}`. A `404` is treated as success, since the message being
already gone is the desired end state.

## Development

```
cd outlook-plugin
npm install
npm run dev-server     # serves https://localhost:3000 with HMR
npm start              # sideloads manifest.xml into Outlook and starts debugging
npm stop               # ends the debugging session
```

`npm start` uses `office-addin-debugging`, which needs the dev certificates trusted on first run
(handled by `office-addin-dev-certs`). The `dev_server_port` in `package.json` (`3000`) must match
the URLs in `manifest.xml`.

Other useful scripts:

```
npm run validate       # validate manifest.xml
npm run build          # production bundle into dist/
npm run build:dev      # development bundle
npm run watch          # rebuild on change
npm run deploy         # publish dist/ to GitHub Pages via gh-pages
```

`npm run lint` currently fails: the repo has an `.eslintrc.json` while the installed ESLint is v9,
which requires the flat `eslint.config.js` format. Unrelated to the add-in logic.

## Deployment

The production build rewrites every `https://localhost:3000/` occurrence to
`https://painavin.github.io/junkbegone/` (including inside `manifest.xml`), then `npm run deploy`
pushes `dist/` to the `gh-pages` branch. Distribute the *rewritten* `dist/manifest.xml` to
sideload against production — not the repo-root `manifest.xml`, which still points at localhost.
