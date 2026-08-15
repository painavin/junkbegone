# Outlook Plugin

An Office.js mail add-in that scans the **Junk Email** folder and deletes messages you don't want,
either on demand (Preview / Run Cleanup) from a task pane.

Source lives in [`outlook-plugin/`](../outlook-plugin).

## What it does

The add-in shows a task pane with an editable bad-word list and two buttons:

| Button | Behaviour |
| --- | --- |
| **Preview** | Scans Junk Email and lists what *would* be deleted, with the reason for each match. Deletes nothing. |
| **Run Cleanup** | Scans Junk Email and permanently deletes every matching message. |

Always use **Preview** first after editing the list — deletion is not undoable through the add-in.

## Deletion rules

A junk message is deleted if **either** rule matches:

1. **Sender matches the bad-word list** — each entry is tested against the string
   `"<sender display name> <sender address>"`.
   - A plain entry (`Americor`) is a case-insensitive substring match.
   - An entry wrapped in slashes (`/Credit.*Card/`, `/Cloud.*Stor/i`) is compiled as a JavaScript
     regex, with optional trailing flags. A malformed regex falls back to substring matching rather
     than throwing.
2. **The message carries an active follow-up flag** — i.e. Graph reports
   `flag.flagStatus === "flagged"`.

Rule 2 lets you flag junk by hand in Outlook and have the next run delete it, without adding a
pattern to the list. Only an *active* flag counts; `complete` (a flag you've already ticked off) and
`notFlagged` are ignored.

> **Careful:** a follow-up flag conventionally means "keep this, follow up on it". Here it means the
> opposite. Do not use flags to rescue a false positive sitting in Junk — flagging it marks it for
> deletion. Move it out of the Junk folder instead.

Because a flagged message matches regardless of who sent it, matches may have no sender at all; the
preview renders those as `(no sender)`.

Relevant code in [`src/taskpane/taskpane.js`](../outlook-plugin/src/taskpane/taskpane.js):
`senderMatchesBadWord`, `isFlagged`, `shouldDelete`, `matchReason`.

## The bad-word list

The list is stored as `junk-senders.json` in Blob Storage — the same blob the Azure Function reads.
It is the single source of truth: editing it in the task pane changes what the scheduled cleanup does,
and **Preview** shows exactly what the scheduled cleanup would delete. See
[azure-function.md](./azure-function.md) for the blob's location and format.

The add-in reads the blob over the Blob REST API with plain `fetch` (no `@azure/storage-blob`, to keep
the bundle small) and writes it back on **Save List**. `DEFAULT_BAD_WORDS` is only a fallback for when
the blob does not exist yet.

A local copy is kept at [`az-function/junk-senders.json`](../az-function/junk-senders.json) for
git-diffable editing, but it is only a copy — the blob is what both components read.

### Loading and the sign-in constraint

`Office.onReady` fires without a user gesture, so a sign-in popup there would be blocked by the
browser. `populateBadWords()` therefore requests a **silent-only** token: if one is cached the
textarea is filled from the blob, otherwise the pane shows *"Sign in with Preview or Run Cleanup to
load the saved list."* Preview and Run Cleanup run from a real click, so they may sign in
interactively, and they fill the textarea if it hasn't loaded yet.

### Guards against losing the list

Two failure modes are handled explicitly, because both silently destroy the list:

- **Saving before loading.** If the blob hasn't loaded, the textarea holds placeholder text rather
  than the real list, so saving would overwrite it. `saveBadWordsFromTextarea()` refuses, and says to
  run Preview first. `badWordsLoaded` tracks this.
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
| Microsoft Graph | `Mail.ReadWrite`, `User.Read` | reading and deleting junk mail |
| Azure Storage | `https://storage.azure.com/user_impersonation` | reading and writing the list blob |

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
