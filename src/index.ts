import express, { NextFunction } from "express";
import session from "express-session";
import morgan from "morgan";
import cookieParser from "cookie-parser";
import { config } from "dotenv";
import bodyParser from "body-parser";
import debug from "debug";
import fileUpload from "express-fileupload";
import flash from "express-flash";
import * as msal from "@azure/msal-node";
import auth from "./auth";
import importSkills from "./skills";

// Load the environment variables
config();

// Set up logging
const log = debug("app:server");

// Set up express
const app = express();
app.use(morgan("tiny"));
app.use(cookieParser());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.json());
app.set("view engine", "ejs");
app.use(fileUpload());
app.use(session({ secret: process.env.SESSION_SECRET as string, resave: true, saveUninitialized: true }));
app.use(flash());

// Local in-memory cache for users.
app.locals.users = {};

// MSAL config
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID as string,
        authority: process.env.AUTHORITY as string,
        clientSecret: process.env.CLIENT_SECRET as string
    },
    system: {
        loggerOptions: {
            loggerCallback(_logLevel: any, message: any, _containsPii: any) {
                debug("app:msal")(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Error
        }
    }
};

// Create msal application object
app.locals.msalClient = new msal.ConfidentialClientApplication(msalConfig);

// Set up routes

// Default rout
app.get("/", (req, res) => {
    log(req.session.userId || "No user");
    const user = req.session.userId ? req.app.locals.users[req.session.userId] : undefined;
    log(user);
    res.render("home", { user });
});

// Auth routes
app.use('/auth', auth);

// Helper method for checking signed in status
const checkSignedIn = (req: any, res: any, next: NextFunction) => {
    const user = req.session.userId ? req.app.locals.users[req.session.userId] : undefined;
    if (user) return next();
    else return res.status(401).send("Not signed in");
}

// upload endpoint
app.post("/upload",
    checkSignedIn,
    async (req, res) => {
        if (!req.files) {
            return res.status(400).send("No files were uploaded.");
        }
        const file = req.files.file as fileUpload.UploadedFile;
        const data = file.data.toString("utf8");
        log(data);
        await importSkills(data, req.app.locals.msalClient, req.session.userId as string);
        res.redirect("/");
    });


// run the app
const port = process.env.PORT || 3210;
app.listen(port, () => {
    log("Listening on port " + port);
});