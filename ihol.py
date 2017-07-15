#!/usr/bin/env python3

import argparse
import sys
import time

from datetime import datetime
from dateutil import tz
from bs4 import BeautifulSoup

import O365 as o365


def get_args(args):
    """ Setup and parse options.
    """
    parser = argparse.ArgumentParser(
            description='I Hate OutLook',
            epilog='** Password read from STDIN.')
    parser.add_argument('email', type=str, help='Email address of user.')
    parser.add_argument('-r', '--remind', action='store_true', default=True,
            help='Remind formatted calendar entries to stdout.')
    parser.add_argument('-b', '--bodies', action='store_true',
            help='Print bodies of calendar to stdout.')
    parser.add_argument('-n', '--num-bodies', type=int, default=3, metavar='N',
            help='Output bodies for N days (default 3).')
    return parser.parse_args(args)

def read_pass(stdin=sys.stdin):
    """ Read password from stdin.
    """
    return stdin.readline()

def event2Remind(ev):
    rem = ['REM']
    start = utc2local(ev.getStart())
    rem.append(start.strftime("%b %d %Y AT %H:%M"))
    rem.append("%%\"%s%%\"" % ev.getSubject())
    return " ".join(rem)

def remindOut(cal):
    cal.getEvents()
    for e in cal.events:
        print(event2Remind(e))

def showBodies(cal, args):
    five_days = time.gmtime(time.time() + (3600*24*args.bodies))
    cal.getEvents(end=time.strftime(cal.time_string, five_days))
    for e in reversed(cal.events):
        if "DevOps Standup" in e.getSubject(): continue
        print("-"*70)
        print("\n".join(formatBody(e)))

fg = "\x1b[38;5;%dm"
hs = fg % 9
hd = fg % 7
cl = "\x1b[0m"
def formatBody(ev):
    soup = BeautifulSoup(ev.getBody(), "html.parser")
    paras = [p.get_text().strip() for p in soup.find_all("p")]
    start = utc2local(ev.getStart()).strftime("%b %d %Y AT %H:%M")
    return ["%s-- %s --%s" % (hs, ev.getSubject(), cl),
            "%s[%s]%s\n" % (hd, start, cl), *["%s\n" % p for p in paras if p]]

def utc2local(t_st):
    """ Times are UTC but timezone isn't set, so we need to set and convert.
    """
    secs = time.mktime(t_st)
    dt = datetime.fromtimestamp(secs)
    dt = dt.replace(tzinfo=tz.tzutc())
    return dt.astimezone(tz.tzlocal())

def main():
    args = get_args(sys.argv[1:])
    passwd = read_pass()
    s = o365.Schedule((args.email, passwd))
    s.getCalendars()
    cal = s.calendars[0]
    if args.bodies:
        showBodies(cal, args)
    else:
        remindOut(cal)


if __name__ == '__main__':
    main()
