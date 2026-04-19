# Spin & Win System (HTML + CSS + JS + Google Sheets)

This project now includes:

- Responsive Spin & Win wheel
- Tab-based layout (`Customer Spin`, `Settings Panel`, `Prizes`) with mobile hamburger menu
- Customer input + one-spin-per-entry protection
- Weighted prize probabilities (supports decimals and `0%`)
- Coupon/receipt image generation (download + WhatsApp share)
- Google Sheets backend with create, list, update, delete sync
- Google Sheets settings sync (shop name, expiry, manual time, app URL, and prize setup)
- Full **Prizes** admin section with search, sort/filter, edit, status update, and delete

## Project Files

- `index.html` - Main UI
- `style.css` - Styling + responsive dashboard
- `script.js` - Frontend logic + sync API calls
- `apps_script.gs` - Google Apps Script backend API

## 1. Run Frontend

Recommended:

```bash
python -m http.server 5500
```

Open:

```text
http://localhost:5500
```

You can also open `index.html` directly, but local server is better for API testing.

## 2. Setup Google Sheet + Apps Script

1. Create a Google Sheet (example name: `SpinWin`).
2. Open `Extensions -> Apps Script`.
3. Paste code from `apps_script.gs`.
4. Save.

## 3. Deploy Apps Script Web App

1. Click `Deploy -> New deployment`.
2. Deployment type: `Web app`.
3. Execute as: `Me`.
4. Who has access: `Anyone`.
5. Deploy and authorize.
6. Copy the `Web app URL`.

## 4. Connect Frontend to Backend

1. Open app UI.
2. In **Settings Panel**, paste URL in `Google Apps Script Web App URL`.
3. Click `Save Settings`.
4. Optional: click `Sync From Google Sheets` in **Prizes** section.

## 5. How Record Sync Works

Backend supports:

- `GET ?action=list` -> load prize records
- `GET ?action=settings` -> load app settings and prize setup
- `POST` -> create record
- `POST { action: "save_settings", settings: {...} }` -> save app settings and prize setup
- `PUT` -> update record
- `DELETE` -> delete record

Because Apps Script Web Apps may not accept raw `PUT/DELETE` in some environments, frontend also supports method override fallback (`POST` + `_method`).
For GET/listing, frontend also supports JSONP fallback when strict CORS blocks normal fetch.

## 6. Admin Prizes Section Features

- Search by unique ID (partial/instant)
- Filter/sort:
  - Recent -> Old
  - Old -> Recent
  - Big -> Small amount
  - Small -> Big amount
  - Completed / Pending / Rejected / Expired
- Actions per record:
  - Edit (customer, amount, prize, expiry)
  - Mark Completed
  - Mark Rejected
  - Delete (confirmation popup)
- Auto status logic:
  - Default = Pending
  - If expiry passed and still pending -> Expired
  - Manual override supports Completed/Rejected

## 7. Hosting

Host these frontend files on any static host:

- `index.html`
- `style.css`
- `script.js`

Examples:

- GitHub Pages
- Netlify
- Vercel
- Shared hosting

Keep Apps Script deployed separately and use its URL in settings.
