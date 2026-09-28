# Junk Be Gone — Documentation

Two independent components clear unwanted mail out of the Outlook **Junk Email** folder using the
same matching rules:

- **[Outlook plugin](./outlook-plugin.md)** — an Office.js task pane for on-demand cleanup, with a
  Preview mode that shows what would be moved before anything is.
- **[Azure Function](./azure-function.md)** — a timer trigger that does the same thing unattended
  every 12 hours.

Both move a junk message to Deleted Items, marked as read with its flag cleared, when the sender matches a bad-word
pattern, the message carries an active follow-up flag, or no word of the sender's display name
appears in their address. A matched message with an emoji in its subject or sender name is
permanently deleted instead. They share an app registration
(`ed7fafe8-a6fc-4ca8-ae75-db77f33f2c5f`), delegated Graph scopes (`Mail.ReadWrite`, `User.Read`), and
the bad-word list itself — stored as `junk-senders.json` in Blob Storage, edited from the add-in's
task pane and read by both. The matching *logic*, however, is duplicated in both codebases; change one
and change the other.
