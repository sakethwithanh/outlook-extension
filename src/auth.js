import { CLIENT_ID, AUTHORITY_TENANT, GRAPH_SCOPES } from "./config.js";

let pcaPromise = null;

// Nested App Authentication. Inside Outlook the host brokers the token, so the
// user is not prompted again after the first consent. Outside Outlook (plain
// browser) MSAL falls back to a normal popup, which keeps local testing easy.
function getClient() {
  if (!pcaPromise) {
    pcaPromise = msal.createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${AUTHORITY_TENANT}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: "localStorage" },
    });
  }
  return pcaPromise;
}

export async function getGraphToken() {
  const pca = await getClient();
  const request = { scopes: GRAPH_SCOPES };

  const accounts = pca.getAllAccounts();
  if (accounts.length > 0) request.account = accounts[0];

  try {
    const r = await pca.acquireTokenSilent(request);
    return r.accessToken;
  } catch (e) {
    // First run, expired session, or consent required.
    const r = await pca.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return r.accessToken;
  }
}
