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

function getContentTypes() {
  var s = SettingsService.getAll();
  return DotCMSApi.getContentTypes(s.hostUrl, s.apiToken);
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
  return fields;
}

function hasExistingMetadataTable() {
  return DocParser.findMetadataTable() !== null;
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
