#!/usr/bin/env python

import O365 as o365
import os

# O365 logging
#import logging
#logging.basicConfig(level=logging.DEBUG)

# Running this triggers a manual copy-paste into browser, login, copy-paste
# back. It should only need to be done once. See README for more.

creds = (os.getenv('IHOL_CLIENT_ID'), os.getenv('IHOL_CLIENT_SECRET'))
acct = o365.Account(credentials=creds)
scopes = ['calendar', 'basic']
result = acct.authenticate(scopes=scopes)
