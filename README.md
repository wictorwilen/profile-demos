# Demos of profile, person or user scenarios in Microsoft Graph

This is a (set of) demo(s) of how to use the profile, person and/or user entities in Microsoft Graph.

## Current demos

* Sample application on how to upload a CSV file with skills for users

## How to run

1. Create a file called `.env` with the contents below
2. Create an Entra ID App with the `User.ReadWrite` user scope and `User.ReadWrite.All` app scope
3. Run the demo with `npm start`

## `.env` file

``` ini
CLIENT_ID=<insert Entra ID app Id>
CLIENT_SECRET=<insert client secret>
SCOPES=User.ReadWrite
AUTHORITY=https://login.microsoftonline.com/common/
REDIRECT_URL=http://localhost:3210/auth/callback
SESSION_SECRET=mysecret
DEBUG=app:*
```