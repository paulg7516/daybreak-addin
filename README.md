# Daybreak Outlook compose add-in

Tags an outgoing message with `X-PTO-Triage` so it lands in the recipient's Daybreak
queue. Static task pane + Office.js, no build, no server.

## Dev: sideload against your own mailbox

1. Trust a localhost HTTPS cert (once):
   `npx -y office-addin-dev-certs install`
2. Serve the `addin/` folder over HTTPS on port 3000:
   `npx -y http-server addin -p 3000 -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key`
   (Path to the cert/key is printed by step 1.)
3. Sideload `addin/manifest.xml`:
   - Outlook on the web: Settings -> Mail -> Customize actions / Get Add-ins -> My add-ins -> Add a custom add-in -> Add from file -> pick `manifest.xml`.
   - Or: `npx -y office-addin-debugging start addin/manifest.xml`

## Verify (manual)

- Compose a new message TO yourself. Open the Daybreak button in the compose ribbon.
- Pick "Action needed", choose a date, leave bcc empty, click Apply -> status shows "Tagged. Send when ready."
- Send it. In the received copy, view the message source/headers and confirm `X-PTO-Triage: action;by=YYYY-MM-DD` is present.
- Run the recipient side: `npm run ingest -- --since <a date before the send>` (or the desktop app) and confirm the message is tagged/placed by the rule.
- Repeat with a bcc address: confirm the bcc recipient receives the message and the header.

## Deploy: GitHub Pages + M365 admin center

1. Push the repo to GitHub and enable Pages serving the `addin/` path (Settings -> Pages),
   giving a base like `https://<you>.github.io/<repo>/addin`.
2. In `manifest.xml`, replace every `https://localhost:3000` with that base, then re-validate:
   `npx -y office-addin-manifest validate addin/manifest.xml`.
3. Microsoft 365 admin center -> Settings -> Integrated apps -> Upload custom apps -> upload `manifest.xml`,
   and assign it to the team. (Free.)
