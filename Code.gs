/* ========= Writing Coach (Google Docs Add-on) =========
   Open-source friendly: no secrets in code.
   Users paste their own API key and backend URL in Settings.
*/

// Store user properties (per-user settings, not global)
var USER_PROP = PropertiesService.getUserProperties();
var KEY_NAME = 'WRITING_COACH_API_KEY';   // Key name for API key
var API_URL_NAME = 'WRITING_COACH_API_URL'; // Key name for API URL

// Run on install and also on open
function onInstall(e) {
  onOpen(e);
}

// Create the custom "Writing Coach" menu in Google Docs
function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Writing Coach')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Set API Key / API URL', 'showSettings')
    .addToUi();
}

// Show the sidebar (loads Sidebar.html)
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Writing Coach');
  DocumentApp.getUi().showSidebar(html);
}

// Show the settings dialog (loads Settings.html)
function showSettings() {
  var html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(420)
    .setHeight(320);
  DocumentApp.getUi().showModalDialog(html, 'Writing Coach Settings');
}

// Save API key and URL into User Properties (private to each user)
function saveSettings(input) {
  var apiKey = input.apiKey;
  var apiUrl = input.apiUrl;

  if (typeof apiKey === 'string') {
    USER_PROP.setProperty(KEY_NAME, apiKey.trim());
  }
  if (typeof apiUrl === 'string') {
    USER_PROP.setProperty(API_URL_NAME, apiUrl.trim());
  }
  return { ok: true };
}

// Get currently selected text in the Google Doc
function getSelectedText() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  if (!selection) return '';
  var elements = selection.getRangeElements();
  var out = [];

  for (var i = 0; i < elements.length; i++) {
    var el = elements[i];
    if (el.isPartial()) {
      // If user only selected part of a paragraph
      out.push(el.getElement().asText().getText().substring(el.getStartOffset(), el.getEndOffsetInclusive() + 1));
    } else {
      // If whole paragraph/element is selected
      var elem = el.getElement();
      if (elem.editAsText) out.push(elem.asText().getText());
    }
  }
  return out.join('\n').trim();
}

// Call external API (backend or OpenAI/Gemini) with selected text + prompt
function analyzeText(input) {
  var apiKey = USER_PROP.getProperty(KEY_NAME) || '';
  var apiUrl = USER_PROP.getProperty(API_URL_NAME) || '';
  if (!apiKey || !apiUrl) {
    throw new Error('Missing API key or API URL. Open "Set API Key / API URL" first.');
  }

  // Build request body
  var payload = {
    prompt: input.prompt || 'Improve clarity and tone. Keep meaning.',
    text: input.text || '',
    audience: input.audience || 'General',
    style: input.style || 'Concise'
  };

  // Send request
  var res = UrlFetchApp.fetch(apiUrl, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload)
  });

  // Handle errors
  var status = res.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Backend error ' + status + ': ' + res.getContentText());
  }

  // Parse response
  var data = JSON.parse(res.getContentText());
  return { suggestion: data.suggestion || '' };
}

// Insert suggestion into the Doc at cursor or end
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
