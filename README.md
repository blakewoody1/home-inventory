# Home Inventory

A home-screen web app for the "Home Inventory" Google Sheet: appliances and electronics with photos, serial numbers, receipts, manuals, and warranty dates.

- `index.html` is the whole app.
- `apps-script/Code.gs` is the backend. Paste it into the sheet's Apps Script editor, set `KEY`, and deploy it as a web app (Execute as: Me, Access: Anyone).
- In the app, tap the gear and enter the web app URL and key. They are saved only on that device.
- Bump `VERSION` in `sw.js` whenever `index.html` changes.
