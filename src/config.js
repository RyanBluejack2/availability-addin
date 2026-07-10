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
  CLIENT_ID: "3e87ee62-ca64-4add-bd22-4030e3eeb25e",
  TENANT_ID: "cda4b246-a751-401b-a7b5-22fece47f850",
  REDIRECT_URI: "https://ryanbluejack2.github.io/availability-addin/index.html",

  // Delegated scope needed to read the signed-in user's own calendar.
  SCOPES: ["Calendars.Read"]
};
