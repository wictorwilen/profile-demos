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


function getAuthenticatedClient(msalClient: msal.ConfidentialClientApplication, userId: string) {
    if (!msalClient || !userId) {
        throw new Error(
            `Invalid MSAL state. Client: ${msalClient ? 'present' : 'missing'}, User ID: ${userId ? 'present' : 'missing'}`);
    }

    // Initialize Graph client
    const client = graph.Client.init({
        // Implement an auth provider that gets a token
        // from the app's MSAL instance
        authProvider: async (done) => {
            try {
                // Get the user's account
                const account = await msalClient
                    .getTokenCache()
                    .getAccountByHomeId(userId);

                if (account) {
                    // Attempt to get the token silently
                    // This method uses the token cache and
                    // refreshes expired tokens as needed
                    const response = await msalClient.acquireTokenSilent({
                        scopes: (process.env.SCOPES as string).split(','),
                        account: account
                    });
                    if(response == null) {
                        throw "No Response";
                    }
 
                    // First param to callback is the error,
                    // Set to null in success case
                    done(null, response.accessToken);
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