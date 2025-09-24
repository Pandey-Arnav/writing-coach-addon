/* ========= Writing Coach (Google Docs Add-on) =========
   This version includes:
   - Menu + sidebar for text improvement
   - Settings dialog for API Key and API URL
   - Logging every request into a Google Sheet
   -----------------------------------------------------
   HOW LOGGING WORKS:
   • You need to create a Google Sheet manually.
   • Copy its URL and paste it into the SETTINGS dialog.
   • Each time you run “Get suggestion,” a row is added:
       Timestamp | Prompt | Input Text | Suggestion
*/

// Keys for storing settings in User Properties
var USER_PROP = PropertiesService.getUserProperties();
var KEY_NAME = 'WRITING_COACH_API_KEY';      // User’s API key
var API_URL_NAME = 'WRITING_COACH_API_URL'; // Backend endpoint
var SHEET_URL_NAME = 'WRITING_COACH_SHEET'; // Google Sheet URL

// Run when installed
function onInstall(e) {
  onOpen(e);
}

// Add menu to Google Docs
function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Writing Coach')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Set Settings', 'showSettings')
    .addToUi();
}

// Open sidebar UI (Sidebar.html)
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Writing Coach');
  DocumentApp.getUi().showSidebar(html);
}

// Open settings dialog (Settings.html)
function showSettings() {
  var html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(420)
    .setHeight(360);
  DocumentApp.getUi().showModalDialog(html, 'Writing Coach Settings');
}

// Save settings from the dialog into User Properties
function saveSettings(input) {
  var apiKey = input.apiKey;
  var apiUrl = input.apiUrl;
  var sheetUrl = input.sheetUrl;

  if (typeof apiKey === 'string') USER_PROP.setProperty(KEY_NAME, apiKey.trim());
  if (typeof apiUrl === 'string') USER_PROP.setProperty(API_URL_NAME, apiUrl.trim());
  if (typeof sheetUrl === 'string') USER_PROP.setProperty(SHEET_URL_NAME, sheetUrl.trim());

  return { ok: true };
}

// Get the currently selected text in the Doc
function getSelectedText() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  if (!selection) return '';
  var elements = selection.getRangeElements();
  var out = [];

  for (var i = 0; i < elements.length; i++) {
    var el = elements[i];
    if (el.isPartial()) {
      // If part of a paragraph is selected
      out.push(el.getElement().asText().getText().substring(el.getStartOffset(), el.getEndOffsetInclusive() + 1));
    } else {
      // If whole paragraph is selected
      var elem = el.getElement();
      if (elem.editAsText) out.push(elem.asText().getText());
    }
  }
  return out.join('\n').trim();
}

// Analyze text by calling backend API
function analyzeText(input) {
  var apiKey = USER_PROP.getProperty(KEY_NAME) || '';
  var apiUrl = USER_PROP.getProperty(API_URL_NAME) || '';
  var sheetUrl = USER_PROP.getProperty(SHEET_URL_NAME) || '';

  if (!apiKey || !apiUrl) {
    throw new Error('Missing API key or API URL. Set them in Settings.');
  }

  // Build request payload
  var payload = {
    prompt: input.prompt || 'Improve clarity and tone. Keep meaning.',
    text: input.text || '',
    audience: input.audience || 'General',
    style: input.style || 'Concise'
  };

  // Call external API
  var res = UrlFetchApp.fetch(apiUrl, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload)
  });

  var status = res.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Backend error ' + status + ': ' + res.getContentText());
  }

  var data = JSON.parse(res.getContentText());
  var suggestion = data.suggestion || '';

  // Log into Google Sheet (if configured)
  if (sheetUrl) {
    try {
      var sheet = SpreadsheetApp.openByUrl(sheetUrl).getActiveSheet();
      sheet.appendRow([
        new Date(),
        payload.prompt,
        payload.text,
        suggestion
      ]);
    } catch (err) {
      // Fail silently if sheet logging fails
      Logger.log("Logging failed: " + err.message);
    }
  }

  return { suggestion: suggestion };
}

// Insert suggestion back into the Doc
function insertSuggestion(text) {
  if (!text) return { ok: true };
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();
  var body = doc.getBody();

  if (cursor) {
    var elem = cursor.insertText(text + '\n');
    if (!elem) body.appendParagraph(text);
  } else {
    body.appendParagraph(text);
  }
  return { ok: true };
}
