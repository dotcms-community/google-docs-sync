/**
 * Per-user settings stored in UserProperties.
 */
var SettingsService = {
  PROP_HOST: 'dotcms_host_url',
  PROP_TOKEN: 'dotcms_api_token',

  getAll: function () {
    var props = PropertiesService.getUserProperties();
    return {
      hostUrl: (props.getProperty(this.PROP_HOST) || '').replace(/\/+$/, ''),
      apiToken: props.getProperty(this.PROP_TOKEN) || ''
    };
  },

  save: function (hostUrl, apiToken) {
    var props = PropertiesService.getUserProperties();
    props.setProperty(this.PROP_HOST, (hostUrl || '').replace(/\/+$/, ''));
    props.setProperty(this.PROP_TOKEN, apiToken || '');
  },

  isConfigured: function () {
    var s = this.getAll();
    return s.hostUrl !== '' && s.apiToken !== '';
  }
};
