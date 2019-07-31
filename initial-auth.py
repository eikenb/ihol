#!/usr/bin/env python

import O365 as o365
import json

# O365 logging
#import logging
#logging.basicConfig(level=logging.DEBUG)

# Running this triggers a manual copy-paste into browser, login, copy-paste
# back. It should only need to be done once. See README for more.

with open("conf.json","r") as f:
    data=json.load(f)

creds = (data['ClientID'], data['ClientSecret'])
acct = o365.Account(credentials=creds)
scopes = ['calendar', 'basic']
result = acct.authenticate(scopes=scopes,tenant_id=data['TenantID'])
