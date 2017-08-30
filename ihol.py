#!/usr/bin/env python3

import argparse
import sys
import time

from datetime import datetime
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
            epilog='** Password read from STDIN.')
    parser.add_argument('email', type=str, help='Email address of user.')
    parser.add_argument('-n', '--next', action='store_true',
            help='Output bodies for next event.')
    output = parser.add_mutually_exclusive_group()
    output.add_argument('-b', '--bodies', action='count', default=1,
            help='Print event bodies for today to stdout (+1 day for each b).')
    output.add_argument('-i', '--ical', action='store_true',
            help='Ical formatted entries to stdout.')
    output.add_argument('-r', '--remind', action='store_true',
            help='Remind formatted calendar entries to stdout.')
    return parser.parse_args(args)

def read_pass(stdin=sys.stdin):
    """ Read password from stdin.
    """
    return stdin.readline()

def event2Remind(ev):
    rem = ['REM']
    start = utc2local(ev.getStart())
    rem.append(start.strftime("%b %d %Y AT %H:%M +10"))
    rem.append("%%\"%s%%\"" % ev.getSubject())
    return " ".join(rem)

def remindOut(cal):
    cal.getEvents()
    for e in cal.events:
        print(event2Remind(e))

def event2Ical(ev):
    c = ical.Calendar()
    c.add('version', '2.0')
    e = ical.Event()
    e.add('dtstart', utc2local(ev.getStart()))
    e.add('dtend', utc2local(ev.getEnd()))
    e.add('summary', ev.getSubject())
    e.add('description', "\n".join(paragraphs(ev)))
    c.add_component(e)
    return c.to_ical().decode("utf-8")

def icalOut(cal):
    cal.getEvents()
    for e in cal.events:
        print(event2Ical(e))
        break

def showBodies(cal, args, next_only=False):
    n = 1 if next_only else args.bodies
    n_days = time.gmtime(time.time() + (3600*24*n))
    cal.getEvents(
            start=time.strftime(cal.time_string, time.gmtime(time.time())),
            end=time.strftime(cal.time_string, n_days))
    for e in reversed(cal.events):
        if e.getSubject() in skip_subjects: continue
        print("-"*70)
        print("\n".join(formatBody(e)))
        if next_only: return

fg = "\x1b[38;5;%dm"
hs = fg % 9
hd = fg % 7
cl = "\x1b[0m"
def formatBody(ev):
    paras = paragraphs(ev)
    start = utc2local(ev.getStart()).strftime("%b %d %Y AT %H:%M")
    return ["%s-- %s --%s" % (hs, ev.getSubject(), cl),
            "%s[%s]%s\n" % (hd, start, cl), *["%s\n" % p for p in paras if p]]

def paragraphs(ev):
    """ Return list of text paragraphs in body of event.
    """
    soup = BeautifulSoup(ev.getBody(), "html.parser")
    return [p.get_text().strip() for p in soup.find_all("p")]

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
    if args.remind:
        remindOut(cal)
    elif args.ical:
        icalOut(cal)
    elif args.next:
        showBodies(cal, args, next_only=True)
    else:
        showBodies(cal, args)


if __name__ == '__main__':
    main()
