/*
 * Optional SharePoint connection defaults.
 *
 * Leave clientId blank to run the app in upload-only mode - everything except the
 * "SharePoint links" tab works without any Azure setup at all.
 *
 * Filling this in bakes the connection into the deployment so your team does not
 * have to enter it themselves. Anything entered under Settings in the app overrides
 * these values for that browser only.
 *
 * These are not secrets. A public-client app registration has no client secret, and
 * every user still has to sign in with their own Microsoft 365 account and can only
 * reach files they already have permission to open.
 */
window.MERGER_CONFIG = {
  clientId: '',
  tenantId: 'common',
};
