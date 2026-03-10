/**
 * Google Docs → dotCMS Sync Add-on
 * Main entry point: menu creation and sidebar launcher.
 */

function onOpen() {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/**
 * Workspace Add-on homepage trigger — returns a card with an "Open Sidebar" button.
 */
function onHomepage() {
  var action = CardService.newAction().setFunctionName('openSidebarAction');
  var button = CardService.newTextButton()
    .setText('Open dotCMS Sync')
    .setOnClickAction(action);
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText('Sync this Google Doc to dotCMS.'))
    .addWidget(button);
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('dotCMS Sync'))
    .addSection(section)
    .build();
}

function openSidebarAction() {
  showSidebar();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('dotCMS Sync')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

// ── Server-side functions called from the sidebar ──

function getSettings() {
  return SettingsService.getAll();
}

function saveSettings(hostUrl, apiToken) {
  SettingsService.save(hostUrl, apiToken);
}

function searchContentTypes(filter) {
  var s = SettingsService.getAll();
  return DotCMSApi.getContentTypes(s.hostUrl, s.apiToken, filter || '');
}

function getContentTypeFields(contentTypeVar) {
  var s = SettingsService.getAll();
  return DotCMSApi.getContentTypeFields(s.hostUrl, s.apiToken, contentTypeVar);
}

function getSites() {
  var s = SettingsService.getAll();
  return DotCMSApi.getSites(s.hostUrl, s.apiToken);
}

function getFolders(siteId) {
  var s = SettingsService.getAll();
  return DotCMSApi.getFolders(s.hostUrl, s.apiToken, siteId);
}

function getLanguages() {
  var s = SettingsService.getAll();
  return DotCMSApi.getLanguages(s.hostUrl, s.apiToken);
}

function searchContent(query) {
  var s = SettingsService.getAll();
  return DotCMSApi.searchContent(s.hostUrl, s.apiToken, query);
}

function generateMetadataTable(contentTypeVar) {
  var s = SettingsService.getAll();
  var fields = DotCMSApi.getContentTypeFields(s.hostUrl, s.apiToken, contentTypeVar);
  DocParser.generateMetadataTable(fields);
  return fields;  // Return ALL fields so the sidebar can populate body field dropdown
}

function getHostFieldVar(contentTypeVar) {
  var s = SettingsService.getAll();
  var fields = DotCMSApi.getContentTypeFields(s.hostUrl, s.apiToken, contentTypeVar);
  for (var i = 0; i < fields.length; i++) {
    var ft = fields[i].fieldType.toLowerCase().replace(/-/g, '');
    if (ft.indexOf('hostfolder') !== -1) {
      return fields[i].variable;
    }
  }
  return 'host';
}

function addFieldToTable(fieldVariable, fieldInfo) {
  DocParser.addFieldToTable(fieldVariable, fieldInfo);
}

function hasExistingMetadataTable() {
  return DocParser.findMetadataTable() !== null;
}

function getMetadataFields() {
  return DocParser.extractMetadataFields();
}

function updateMetadataField(fieldVariable, value) {
  DocParser.updateMetadataField(fieldVariable, value);
}

function syncToDocCMS(options) {
  return SyncEngine.sync(options);
}

function getSyncLog() {
  return SyncEngine.getSyncLog();
}

function getDebugFireResult() {
  return PropertiesService.getDocumentProperties().getProperty('_debug_fireResult') || 'no data';
}
