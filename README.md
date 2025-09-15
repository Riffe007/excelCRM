# Excel CRM — Excel Add-in (Office.js + D3)

A lightweight **CRM that runs inside Excel (web & desktop)** using **Office.js** + **D3.js**. 
It provides a **task pane app** with a polished dashboard, a **New Lead** form, searchable leads table, and simple activity logging.

---

## ✨ Features
- **D3 Dashboard**: KPIs, Status Mix (pie), Pipeline Value by Stage (bars), and 6‑month new‑leads trend (line).
- **CRM Form**: Owner, Account, Contact info, Source/Priority/Status/Stage, Value, Close Date, Notes.
- **Leads & Activities**: Stored in Excel tables (`LeadsTable`, `ActivitiesTable`) for easy filtering/sorting.
- **Auto‑bootstrap**: Creates sheets/tables if missing.
- **Excel Online Ready**: Pure Office.js (no VBA). Works in **Excel for the web** and **desktop Excel**.

---

## 📁 Repo Structure
```
/web
  ├── index.html      # Task pane UI (tabs for Dashboard, Form, Leads, Activities)
  ├── styles.css      # Dark UI theme
  ├── app.js          # Office.js + D3 logic
  └── icons/
      ├── icon-32.png
      └── icon-80.png
manifest.xml          # Office Add-in manifest 
```

---

## ✅ Requirements
- Microsoft 365 account (to run **Excel for the web**).
- Excel Online (recommended) or Desktop Excel (Microsoft 365 subscription).
- Ability to **Upload a custom Office Add-in** (Excel → Insert → Office Add‑ins → *Upload My Add-in*).

> **No macros/VBA needed.** This is pure HTML/CSS/JS + Office.js.

---

## 🚀 Quick Start (GitHub Pages Deploy)
1. **Host the web app** via GitHub Pages (this repo):
   - Ensure the site is published. The app is expected at:  
     **`https://riffe007.github.io/excelCRM/web/index.html`**
2. **Verify the page opens** in a browser:
   - `https://riffe007.github.io/excelCRM/web/index.html`
3. **Sideload the add-in in Excel for the web**:
   - Open **Excel Online** → any workbook (or `LifeNavigator_CRM_Template.xlsx`).
   - **Insert → Office Add-ins → Upload My Add-in**.
   - Select **`manifest.xml`** (already configured to the URL above).
   - A new **“Open CRM”** button appears on the Home tab; click it to launch the task pane.
4. **First Run**:
   - The add-in will create the `Leads`, `Activities`, and `Accounts` sheets/tables if they don’t exist.
   - Use the **New Lead** form to add your first record. The dashboard updates automatically (use **Refresh** if needed).

---

## 🧭 Using the App
- **Dashboard**: Live KPIs & charts. Click **Refresh** after bulk edits.
- **New Lead**: Submit creates a new row in `LeadsTable` with timestamps and a sequential ID.
- **Leads**: Use the top search to filter by **name/email**. Badges indicate status.
- **Activities**: “Log Activity” prompts for Lead ID and details, writing to `ActivitiesTable`.

> Everything is stored in your workbook, so it’s portable and auditable.

---

## 🛠️ Customization
- **Branding**: Replace `/web/icons/icon-32.png` and `icon-80.png`.  
- **Taxonomies**: Edit `<select>` options in `/web/index.html` (Source, Priority, Status, Stage).
- **Charts**: Tweak D3 visuals in `/web/app.js` (`drawPie`, `drawBars`, `drawLine`).  
- **Data Model**: Extend the headers in `ensureTables()` and the form/table mapping in `onSubmit()`.

If you change any **web paths**, update them in `manifest.xml` (search for `https://.../web/`).

---

## 🧪 Local Development
You can host `/web` locally with HTTPS and update `manifest.xml` to point to your dev URL.
A quick option is to use a tunneling tool (e.g., `localtunnel`, `ngrok`) to expose `https://` for Office.

> Office Add-ins **require HTTPS** for most hosts; `http://localhost` is limited. For production, GitHub Pages is simplest.

---

## 🧩 Excel Desktop
This add-in also works in **desktop Excel**. Sideload via:
- **Insert → Office Add-ins → Upload My Add-in** (if enabled), or
- Centralized deployment / shared catalog (admin-managed).

Excel Online is recommended for first-time setup because **Upload My Add-in** is always available there.

---

## 🔐 Permissions
The manifest requests:
- **ReadWriteDocument** — read/write to the workbook only.  
No external APIs are called by default. If you later add Microsoft Graph or your own API, update the manifest’s
`<WebApplicationInfo>` accordingly.

---

## 🧯 Troubleshooting
- **Add-in fails to load**: Confirm the manifest URLs match your GitHub Pages paths exactly:  
  `https://riffe007.github.io/excelCRM/web/index.html`, `https://riffe007.github.io/excelCRM/web/icons/icon-32.png`, `https://riffe007.github.io/excelCRM/web/icons/icon-80.png`.
- **Blank task pane**: If GitHub Pages just updated, wait ~1–2 minutes and hard-refresh (Ctrl/Cmd+Shift+R).
- **Charts not updating**: Click **Refresh** (top-right). Make sure dates are `YYYY-MM-DD` in `Created On`.
- **Search not filtering**: Search looks at **Name** and **Email** only (adjust in `onSearch()` if needed).
- **No “Upload My Add-in” on desktop**: Use Excel **web** to sideload, or configure a shared catalog / centralized deployment.

---

## 📄 License
MIT © LifeNavigator

---

## 🙋 Support
- Issues: open a GitHub issue in this repo.
- Repo home: https://github.com/Riffe007/excelCRM
