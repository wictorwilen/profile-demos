// add the userId to the express session object
export declare module 'express-session' {
    interface SessionData {
        userId: string;
    }
}

