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

The O365 library used has this documented in their
[README](https://github.com/O365/python-o365#authentication). Do the first step
(with all its sub-steps) to get the `Application ID` and `Client Secret`. When
adding permissions, be sure to add these:

  `Calendars.ReadWrite`, `offline_access` and `User.Read`

Be sure to have the `Application (client) ID` and the `Client Secret` saved. If
there is a `Directory (tenant) ID` in the Overview write it down as well.

**NOTE**: if there is no `Directory (tenant) id`, use "common" TenantID in the
conf.json file.

Once that is done, continue below..

- Copy `conf.json.template` to `conf.json`
- Create your virtualenv buy running `createVenv.sh`
- Populate your `conf.json` with various IDs obtained above
- Run the ./initial-auth.py and follow the directions
- ihol should now run

Note you also need to keep the ./o365_token.txt file created by the
initial-auth.py run. The token in that file will stay good indefinitely as long
as you use it once every 90 days.
