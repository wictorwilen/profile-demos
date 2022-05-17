import * as msal from "@azure/msal-node";
import * as graph from "@microsoft/microsoft-graph-client";
import debug from "debug";
require('isomorphic-fetch');

const log = debug("app:graph");

const getUserDetails = async (msalClient: msal.ConfidentialClientApplication, homeAccountId: string, userId: string) => {
    const client = getAuthenticatedClient(msalClient, homeAccountId);

    const user = await client
        .api(`/users/${userId}`)
        .select('displayName,mail,userPrincipalName')
        .get();
    return user;
}

const getAuthenticatedClient = (msalClient: msal.ConfidentialClientApplication, userId: string, application: boolean = false) => {

    // Initialize Graph client
    const client = graph.Client.init({
        // Implement an auth provider that gets a token
        // from the app's MSAL instance
        authProvider: async (done) => {
            try {
                // Get user info
                const account = await msalClient
                    .getTokenCache()
                    .getAccountByHomeId(userId);

                if (account) {

                    if (application === false) {
                        // Delegated permissions
                        const response = await msalClient.acquireTokenSilent({
                            scopes: (process.env.SCOPES as string).split(','),
                            account: account
                        });
                        if (response == null) {
                            throw "No Response";
                        }

                        // First param to callback is the error,
                        // Set to null in success case
                        done(null, response.accessToken);

                    } else {
                        // Application permissions
                        const response = await msalClient.acquireTokenByClientCredential({
                            scopes: ["https://graph.microsoft.com/.default"],
                            authority: `https://login.microsoftonline.com/${account.tenantId}/`
                        });
                        if (response == null) {
                            throw "No Response";
                        }
                        done(null, response.accessToken);
                    }
                }
            } catch (err) {
                log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
                done(err, null);
            }
        }
    });

    return client;
}

export { getUserDetails, getAuthenticatedClient };