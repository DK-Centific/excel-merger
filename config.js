/*
 * Connection defaults, baked into the deployment so the team does not have to configure
 * anything. Anything entered under Settings in the app overrides these, for that browser only.
 *
 * None of these are secrets:
 *  - A public-client Azure registration has no client secret.
 *  - A Dropbox PKCE app key is public by design; the app *secret* is never used and must
 *    never be put here.
 * Every user still signs in with their own account and can only reach what they could
 * already open.
 *
 * Leave clientId blank to run in upload-only mode; leave dropboxAppKey blank to hide the
 * link filler's Dropbox connection.
 */
window.MERGER_CONFIG = {
  clientId: '',
  tenantId: 'common',

  dropboxAppKey: 'vgmp1z3692hv0rc',

  // Where "Browse Dropbox" opens. Everything below this is selectable.
  dropboxBrowseRoot: '/Centific Team Folder/Agency Collection',
};
