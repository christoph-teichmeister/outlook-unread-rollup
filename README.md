# Outlook web: unread rollup on collapsed folders

Tampermonkey userscript. Outlook on the web shows only a folder's own unread count when it is collapsed. This script shows own + subfolder unread total on collapsed folders, styled like the native badge.

## Install

1. Install [Tampermonkey](https://www.tampermonkey.net/).
2. Open the [raw script](https://raw.githubusercontent.com/christoph-teichmeister/outlook-unread-rollup/main/outlook-unread-rollup.user.js) and click Install. Tampermonkey then updates it from here.

Runs on `outlook.office.com`, `outlook.office365.com` and `outlook.cloud.microsoft` (`/mail/*`).

## How it works

- Collapsed subfolders are not in the page DOM, so the counts come from the API: the script reads the token Outlook web itself keeps in `localStorage` (MSAL cache, scope `outlook.office.com/Mail.*`) and calls the Outlook REST v2.0 folder endpoint. Nothing leaves your browser except requests to `outlook.office.com`.
- Refresh every 90 s and when the tab becomes visible.
- Collapsed rows get the total, expanded rows keep the native count.

## Caveats

- Depends on an undocumented token cache, the legacy REST v2.0 API and Outlook's DOM; it may break after an Outlook update.
- Folders are matched by display name path.
- Work and school (Microsoft 365) accounts only; not tested elsewhere.

## License

MIT
