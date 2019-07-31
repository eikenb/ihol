# I Hate OutLook (ihol)

We use Office-365 at work, but I despise the online Outlook 'client' for
calendaring. So I wrote this. Turns out several of my co-workers are in the
same boat, so I'm sharing this here.

At the moment it outputs calendar events in
[remind](http://www.roaringpenguin.com/products/remind/) or basic
[ICal](https://pypi.python.org/pypi/icalendar) formats for adding them to local
calendars. It will also output the full body of the next or more events.


## Authentication Setup

MS requires Oauth2 based login now which complicates setup a bit. You need a
set of API credentials from MS and store them in a config file in json format.

`conf.js.template` is a template to be renamed to `conf.json` and populate 
with your own id/credentials (`ClientID`, `ClientSecret`, `TenantID`)

The O365 library I use has this documented in their
[README](https://github.com/O365/python-o365#authentication), but the TLDR
version is..

- Copy `conf.json.template` to `conf.json`
- Create your virtualenv buy running `createVenv.sh`
- Login at https://apps.dev.microsoft.com/ with your Office-365 creds.
- Create an app (the application-id is your `ClientID`)
- Generate a client secret in "Certificates & secrets" (this is your `ClientSecret`)
- Under Authentication, add a new web platform and set the redirect URL to: 
    https://login.microsoftonline.com/common/oauth2/nativeclient
- Under the "API Permissions", add new permission : Delegated
    "Calendars.ReadWrite", "offline_access" and "User.Read"
(remember to save)
- Under Overview, Directory (tenant) ID is your `TenantID`
- Populate your `conf.json` with value obtained above
- Run the ./initial-auth.py and follow the directions
- ihol should now run

Note you also need to keep the ./o365_token.txt file
created by the initial-auth.py run. The token in that file will stay good
indefinitely as long as you use it once every 90 days.
