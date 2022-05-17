import express, { Request } from "express";
import passport from "passport";
import session from "express-session";
import { BearerStrategy, IBearerStrategyOptionWithRequest, IOIDCStrategyOptionWithoutRequest, IOIDCStrategyOptionWithRequest, IProfile, OIDCStrategy, VerifyCallback } from "passport-azure-ad";
import morgan from "morgan";
import cookieParser from "cookie-parser";
import { config } from "dotenv";
import bodyParser from "body-parser";
import debug from "debug";
import jwtDecode from "jwt-decode"
import fileUpload from "express-fileupload";

config();

const log = debug("app:server");

// set up auth
const options: IOIDCStrategyOptionWithRequest = {
    identityMetadata: "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
    clientID: process.env.CLIENT_ID as string,
    clientSecret: process.env.CLIENT_SECRET as string,
    responseType: "id_token code",
    responseMode: "form_post",
    redirectUrl: process.env.REDIRECT_URL as string,
    passReqToCallback: true,
    allowHttpForRedirectUrl: true,
    loggingLevel: "error",
    issuer: "https://sts.windows.net/",
    validateIssuer: false
}

passport.serializeUser((user, done) => {
    done(null, user);
});

passport.deserializeUser((user: Express.User | false | null, done) => {
    done(null, user);
});

const strategy = new OIDCStrategy(
    options,
    (req: Request, iss: string, sub: string, profile: IProfile, access_token: string, refresh_token: string, done: VerifyCallback) => {
        log("Verify");
        const decoded: any = jwtDecode(access_token);
        done(null, { ...profile, given_name: decoded.given_name, family_name: decoded.family_name, name: decoded.name, upn: decoded.upn }, access_token)
    });

passport.use(strategy);


// set up express
const app = express();
app.use(morgan("combined"));
app.use(cookieParser());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.json());
app.set("view engine", "ejs");
app.use(fileUpload());

// initialize passport
app.use(session({ secret: process.env.SESSION_SECRET as string, resave: true, saveUninitialized: true }));
app.use(passport.initialize());
app.use(passport.session());


// Set up routes
app.get("/", (req, res) => {
    res.render("home", { user: req.user });
});

app.get('/login',
    passport.authenticate("azuread-openidconnect", { failureRedirect: '/' }),
    (_req, res) => {
        log("Login was called");
        res.redirect('/');
    });

app.post("/_signin",
    passport.authenticate("azuread-openidconnect", {}),
    (req, res) => {

        if (req.body.error) {
            log("Error: " + req.body.error_description);
        } else {
            log(req.body)
        }
        log("Sign in");
        res.redirect('/');
    });

app.get('/logout', function (req, res) {
    req.logout();
    res.redirect('/');
});

const importSkills = async (data: string) => {
    return new Promise<void>((resolve, reject) => {
        data.split("\n").forEach(line => {
            const [name, skill, proficiency, collaborationTags, allowedAudiences] = line.split(";");
            console.log(`${name} knows ${skill}`);
        });
    });
}

app.post("/upload",
    (req, res, next) => {
        if (req.user) return next();
        else return res.status(401).send("Not signed in");
    },
    (req, res) => {
        if (!req.files) {
            return res.status(400).send("No files were uploaded.");
        }
        const file = req.files.file as fileUpload.UploadedFile;
        const data = file.data.toString("utf8");
        log(data);
        importSkills(data);
        res.redirect("/");
    });


// run the app
const port = process.env.PORT || 3210;
app.listen(port, function () {
    log("Listening on port " + port);
});