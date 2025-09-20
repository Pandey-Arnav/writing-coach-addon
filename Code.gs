/* ========= Writing Coach (Google Docs Add-on) =========
   Open-source friendly: no secrets in code.
   Users paste their own API key in the sidebar Settings.
*/

const USER_PROP = PropertiesService.getUserProperties();
const KEY_NAME = 'WRITING_COACH_API_KEY';
const API_URL_NAME = 'WRITING_COACH_API_URL';

function onInstall(e) { onOpen(e); }
function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Writing Coach')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Set API Key / API URL', 'showSettings')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Writing Coach');
  DocumentApp.getUi().showSidebar(html);
}

function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(420)
    .setHeight(320);
  DocumentApp.getUi().showModalDialog(html, 'Writing Coach Settings');
}

function saveSettings({ apiKey, apiUrl }) {
  if (typeof apiKey === 'string') USER_PROP.setProperty(KEY_NAME, apiKey.trim());
  if (typeof apiUrl === 'string') USER_PROP.setProperty(API_URL_NAME, apiUrl.trim());
  return { ok: true };
}

function getSelectedText() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) return '';
  const elements = selection.getRangeElements();
  let out = [];
  elements.forEach(el => {
    if (el.isPartial()) {
      out.push(el.getElement().asText().getText().substring(el.getStartOffset(), el.getEndOffsetInclusive() + 1));
    } else {
      const elem = el.getElement();
      if (elem.editAsText) out.push(elem.asText().getText());
    }
  });
  return out.join('\n').trim();
}

function analyzeText(input) {
  const apiKey = USER_PROP.getProperty(KEY_NAME) || '';
  const apiUrl = USER_PROP.getProperty(API_URL_NAME) || '';
  if (!apiKey || !apiUrl) {
    throw new Error('Missing API key or API URL. Open "Set API Key / API URL" first.');
  }

  const payload = {
    prompt: input.prompt || 'Improve clarity and tone. Keep meaning.',
    text: input.text || '',
    audience: input.audience || 'General',
    style: input.style || 'Concise',
  };

  const res = UrlFetchApp.fetch(apiUrl, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
  });

  const status = res.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Backend error ' + status + ': ' + res.getContentText());
  }

  const data = JSON.parse(res.getContentText());
  return { suggestion: data.suggestion || '' };
}

function insertSuggestion(text) {
  if (!text) return { ok: true };
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const body = doc.getBody();
  if (cursor) {
    const elem = cursor.insertText(text + '\n');
    if (!elem) body.appendParagraph(text);
  } else {
    body.appendParagraph(text);
  }
  return { ok: true };
}
