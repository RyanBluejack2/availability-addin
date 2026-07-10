// ============================================================
//  CONFIG — fill these in after your Azure app registration.
// ============================================================
//
// CLIENT_ID:  the "Application (client) ID" from your Azure app registration.
// TENANT_ID:  your Bluejack directory (tenant) ID, OR the literal string
//             "organizations" to allow any work account in your tenant.
// REDIRECT_URI: must EXACTLY match a redirect URI you register in Azure
//               (Single-page application platform). Use your hosting URL + index.html.
//
window.AVAILABILITY_CONFIG = {
  CLIENT_ID: "REPLACE_WITH_CLIENT_ID",
  TENANT_ID: "REPLACE_WITH_TENANT_ID",
  REDIRECT_URI: "https://REPLACE_WITH_YOUR_HOST/index.html",

  // Delegated scope needed to read the signed-in user's own calendar.
  SCOPES: ["Calendars.Read"]
};
