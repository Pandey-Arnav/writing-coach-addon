# writing-coach-addon
Open-source Google Docs add-on that helps improve writing clarity, tone, and style using your own AI API key.


# Writing Coach (Google Docs Add-on)

Open-source Google Docs add-on that helps improve writing clarity, tone, and style using your own AI API key.

---

## ✨ Features
- Adds a **Writing Coach** menu in Google Docs (under Extensions).
- Sidebar UI for entering text or using the current selection.
- Custom prompts (e.g., “make this more concise”).
- Secure storage of your API key & backend URL in **User Properties**.
- Inserts AI suggestions directly back into your document.
- 100% open-source — no secrets in code.

---

## 🚀 Installation

### Quick Setup (manual copy)
1. Open any Google Doc.
2. Go to **Extensions → Apps Script**.
3. Delete any existing code and copy in:
   - `Code.gs`
   - `Sidebar.html`
   - `Settings.html`
   - `appsscript.json`
4. Save the project.
5. Reload your Doc — you’ll see **Writing Coach** in the menu bar.

### Using `clasp` (Apps Script CLI)
If you want to sync this repo with Google Apps Script:
```bash
npm install -g @google/clasp
clasp login
clasp create --type docs --title "Writing Coach"
clasp push


