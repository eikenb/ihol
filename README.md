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
set of API credentials from MS and store them in environment variables;

The software looks for them in the HOL_CLIENT_ID and IHOL_CLIENT_SECRET. The
O365 library I use has this documented in their
[README](https://github.com/O365/python-o365#authentication), but the TLDR
version is..

0. Create your virtualenv as needed (see ./ihol-venv for guidance)
1. Login at https://apps.dev.microsoft.com/ with your Office-365 creds.
2. Create an app (the application-id is your IHOL_CLIENT_ID)
3. Generate a new password (this is your IHOL_CLIENT_SECRET)
4. Under platform, add a new web platform and set the redirect URL to;
    https://outlook.office365.com/owa/
5. Under the "Microsoft Graph Permissions", delegated permissions add:
    "Calendars.ReadWrite", "offline_access" and "User.Read"
(remember to save)
6. Set the IHOL_CLIENT_ID and IHOL_CLIENT_SECRET environment variables
7. Run the ./initial-auth.py and follow the directions
8. ihol should now run

Note you need the IHOL_CLIENT_ID and IHOL_CLIENT_SECRET set for each run (even
after the initial-auth.py run). You also need to keep the ./o365_token.txt file
created by the initial-auth.py run. The token in that file will stay good
indefinitely as long as you use it once every 90 days.
