import PromiseRouter from "express-promise-router";
import * as graph from "./graph";

const router = PromiseRouter();

// login rout
router.get("/login", async (req, res) => {
    const urlParameters = {
        scopes: (process.env.SCOPES as string).split(','),
        redirectUri: process.env.REDIRECT_URL
    };

    try {
        const authUrl = await req.app.locals.msalClient.getAuthCodeUrl(urlParameters);
        res.redirect(authUrl);
    }
    catch (error) {
        console.log(`Error: ${error}`);
        req.flash('error_msg', 'Error getting auth URL');
        res.redirect('/');
    }
});

// callback route
router.get('/callback',
    async (req, res) => {
        const tokenRequest = {
            code: req.query.code,
            scopes: (process.env.SCOPES as string).split(','),
            redirectUri: process.env.REDIRECT_URL
        };

        try {
            const response = await req.app.locals
                .msalClient.acquireTokenByCode(tokenRequest);

            // Save the user's homeAccountId in their session
            req.session.userId = response.account.homeAccountId;
            if (req.session.userId === undefined) {
                throw "Invalid session";
            }

            const user = await graph.getUserDetails(
                req.app.locals.msalClient,
                req.session.userId,
                response.account.localAccountId
            );

            // Add the user to user storage
            req.app.locals.users[req.session.userId] = {
                displayName: user.displayName,
                email: user.mail || user.userPrincipalName
            };

        } catch (error) {
            req.flash('error_msg', 'Error completing authentication');
        }

        res.redirect('/');
    }
);

// logout route
router.get('/logout',
    async (req, res) => {
        if (req.session.userId) {
            const accounts = await req.app.locals.msalClient
                .getTokenCache()
                .getAllAccounts();

            const userAccount = accounts.find((a: { homeAccountId: any; }) => a.homeAccountId === req.session.userId);

            if (userAccount) {
                req.app.locals.msalClient
                    .getTokenCache()
                    .removeAccount(userAccount);
            }
        }

        req.session.destroy((err) => {
            res.redirect('/');
        });
    }
);

export default router;