#!/usr/bin/env python3

import argparse
import datetime as dt
import os
import re
import sys
import time
import json

from dateutil import tz
from bs4 import BeautifulSoup

import icalendar as ical
import O365 as o365
# https://github.com/Narcolapser/python-o365

# list of event subjects to skip
skip_subjects = ["DevOps Standup"]
# probably should go in a config or something

def get_args(args):
    """ Setup and parse options.
    """
    parser = argparse.ArgumentParser(
            description='I Hate OutLook',
            epilog='** Credentials read from environment variables.')
    parser.add_argument('-n', '--next', action='store_true',
            help='Output bodies for next event.')
    output = parser.add_mutually_exclusive_group()
    output.add_argument('-b', '--bodies', action='count', default=1,
            help='Print event bodies for today to stdout (+1 day for each b).')
    output.add_argument('-i', '--ical', action='store_true',
            help='Ical formatted entries to stdout.')
    output.add_argument('-r', '--remind', action='store_true',
            help='Remind formatted calendar entries to stdout.')
    parser.add_argument('-a', '--append', action='append',
            help='Append to remind output.')
    parser.add_argument('-l', '--limit', action='store', default=60, type=int,
            help='Number of events to fetch.')
    parser.add_argument('-d', '--days', action='store', default=60, type=int,
            help='Days to look ahead for events (for remind/ical output).')
    return parser.parse_args(args)

def get_creds():
    with open("conf.json","r") as f:
        data=json.load(f)
    return data

def event2Remind(ev, args):
    rem = ['REM']
    start = ev.start
    end = ev.end
    duration = end - start
    if duration > dt.timedelta(days=1):
        rem.append(start.strftime("%b %d %Y THROUGH"))
        rem.append(end.strftime("%b %d %Y"))
    else:
        rem.append(start.strftime("%b %d %Y AT %H:%M +10"))
        dur = str(end - start)
        rem.append("DURATION " + dur[:dur.rfind(":")])
    rem.append("MSG %%\"%s%%\"" % ev.subject)
    if args.append: rem.extend(args.append)
    return " ".join(rem)

def calEvents(cal, days, limit):
    n_days = dt.timedelta(days=days)
    q = cal.new_query('start').greater_equal(dt.datetime.now())
    q.chain('and').on_attribute('end').less_equal(dt.datetime.now() + n_days)
    return cal.get_events(limit=limit, query=q)

def remindOut(cal, args):
    events = calEvents(cal, args.days, args.limit)
    for e in events:
        print(event2Remind(e, args))

def icalOut(cal, args):
    events = calEvents(cal, args.days, args.limit)
    for e in events:
        print(event2Ical(e))

def event2Ical(ev):
    c = ical.Calendar()
    c.add('version', '2.0')
    e = ical.Event()
    e.add('dtstart', ev.start)
    e.add('dtend', ev.end)
    e.add('summary', ev.subject)
    e.add('description', bodyText(ev))
    c.add_component(e)
    return c.to_ical().decode("utf-8")

def showBodies(cal, args, next_only=False):
    n = 1 if next_only else args.bodies
    events = calEvents(cal, n, args.limit)
    events = sorted(events, key=lambda e: e.start, reverse=(not next_only))
    for e in events:
        if e.subject in skip_subjects: continue
        print("-"*70)
        print("\n".join(formatBody(e)))
        if next_only: return

fg = "\x1b[38;5;%dm"
hs = fg % 9
hd = fg % 15
cl = "\x1b[0m"
def formatBody(ev):
    start = ev.start.strftime("%b %d %Y AT %H:%M")
    body = ["%s-- %s --%s" % (hs, ev.subject, cl),
            "%s[%s]%s" % (hd, start, cl)]
    if ev.location:
        body.append("%s[%s]%s\n" % (hd, ev.location, cl))
    body.append(bodyText(ev))
    return body

def bodyText(ev):
    """ Return list of text paragraphs in body of event.
    """
    soup = BeautifulSoup(ev.body, "html.parser")
    # strip out head, as comments in it cause issues
    for h in soup.find_all("head"): h.extract()
    #return soup.get_text("\n")
    return re.sub("\\s{2,64}", "\n\n", soup.get_text("\n").strip())

def main():
    args = get_args(sys.argv[1:])
    creds=get_creds()
    credentials = (creds['ClientID'],creds['ClientSecret'])
    scopes = ['calendar', 'basic']
    acct = o365.Account(credentials, scopes=scopes,tenant_id=creds['TenantID'])
    s = acct.schedule()
    cal = s.get_default_calendar()
    if args.remind:
        remindOut(cal, args)
    elif args.ical:
        icalOut(cal, args)
    elif args.next:
        showBodies(cal, args, next_only=True)
    else:
        showBodies(cal, args)


if __name__ == '__main__':
    main()
