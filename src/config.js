// ---------------------------------------------------------------
// Edit these two values, then redeploy. Nothing else needs changing.
// ---------------------------------------------------------------

// Application (client) ID from your Azure app registration.
export const CLIENT_ID = "REPLACE_WITH_AZURE_CLIENT_ID";

// "common"        -> work/school AND personal outlook.com accounts
// "organizations" -> work/school only
// "<tenant-guid>" -> your tenant only (use for the org-wide rollout)
export const AUTHORITY_TENANT = "common";

// Delegated scope. Mail.Read does not require admin consent by default.
export const GRAPH_SCOPES = ["Mail.Read"];

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// Newest message first in the exported file.
export const NEWEST_FIRST = false;

// Max parallel Graph requests when fetching message bodies.
export const FETCH_CONCURRENCY = 4;
