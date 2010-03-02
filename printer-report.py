# Copyright (c) 2010, Aneurin Price <aneurin.price@gmail.com>

# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation
# files (the "Software"), to deal in the Software without
# restriction, including without limitation the rights to use,
# copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following
# conditions:

# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
# OTHER DEALINGS IN THE SOFTWARE.


import smtplib
import win32evtlog
import win32evtlogutil
import winerror
import re
from sys import argv, exit
from datetime import datetime, date, timedelta
from optparse import OptionParser

class Record:
    def addJob(self, pages, bytes):
        self.jobs += 1
        self.pages += pages
        self.bytes += bytes

    jobs = 0
    pages = 0
    bytes = 0

class PrinterRecord(Record):
    def __init__(self):
        self.users = JobDict('User')
        self.groups = JobDict('Group')

class JobDict:
    def __init__(self, name, recordType=Record):
        self.name = name
        self.dict = {}
        self.width = 1
        self.recordType = recordType

    def addJob(self, name, pages, bytes):
        if not name in self.dict:
            self.dict[name] = self.recordType()
            if len(name) > self.width:
                self.width = len(name)
        self.dict[name].addJob(pages, bytes)

    def getPageCount(self, record):
        return record[1].pages

    def summarise(self, printHeader=False):
        lines = []
        if printHeader:
            header = self.name + ' counts:'
            lines.append('%s\n%s\n::\n\n' % (header, '-' * len(header)))
        list = self.dict.items()
        list.sort(key=self.getPageCount, reverse=True)
        for entry in list:
            name = entry[0]
            pages = entry[1].pages
            jobs = entry[1].jobs
            bytes = entry[1].bytes
            lines.append(' %-*s Page count: %-5d\t(jobs: %d,\tpages per job: %.2f,' \
                         '\ttotal print size: %s,\taverage print size: %s)\n' \
                         % (self.width, name, pages, jobs,
                            float(pages) / jobs, size(bytes),
                            size(float(bytes) / jobs))
                        )
        return ''.join(lines)

def dateFromString(dateString):
    dateRegEx = re.compile(r'(?P<year>\d{2,4})(-(?P<month>\d{1,2})(-(?P<day>\d{1,2}))?)?')
    match = dateRegEx.match(dateString)
    if match:
        y, m, d = dateRegEx.match(dateString).group('year', 'month', 'day')
    elif dateString.lower() == 'today':
        today = datetime.today()
        y = today.year
        m = today.month
        d = today.day
    else:
        exit("Error: Couldn't interpret '%s' as a date." % dateString)
    if int(y) < 100:
        y = int(y) + 2000;
    return int(y or 0), int(m or 0), int(d or 0)

def startOf(year, month=0, day=0):
    return datetime(year, month or 1, day or 1)

def endOf(year, month=0, day=0):
    endTime = None
    if month:
        if day:
            endTime = datetime(year, month, day, 23, 59, 59)
        else:
            if month == 12:
                year += 1
                month = 1
            else:
                month += 1
            endTime = datetime(year, month, 1) - timedelta(seconds=1)
    else:
        endTime = datetime(year, 12, 31, 23, 59, 59)
    return endTime

def getStartDate(options):
    year = 0
    month = 0
    day = 0
    if options.timePeriod:
        year, month, day = dateFromString(options.timePeriod)
    elif options.startDate:
        year, month, day = dateFromString(options.startDate)
    else:
        year = datetime.today().year
        month = datetime.today().month - 1
        if month == 0:
            year -= 1
            month = 12
    return startOf(year, month, day)

def getEndDate(options):
    year = 0
    month = 0
    day = 0
    now = datetime.today()
    if options.timePeriod:
        year, month, day = dateFromString(options.timePeriod)
    elif options.startDate:
        year, month, day = dateFromString(options.endDate)
    else:
        year = now.year
        month = now.month - 1
        if month == 0:
            year -= 1
            month = 12
    endDate = endOf(year, month, day)
    if endDate > now:
        endDate = now # This won't be the same as endOf(now.year,...)
    return endDate

def dayAsStr(date):
    suffixes = {}
    suffixes[1] = 'st'
    suffixes[2] = 'nd'
    suffixes[3] = 'rd'
    suffixes[21] = 'st'
    suffixes[22] = 'nd'
    suffixes[23] = 'rd'
    suffixes[31] = 'st'
    return str(date.day) + suffixes.get(date.day, 'th')

def getTimePeriodAsStr(start, end):
    timePeriodStr = None
    if start.year == end.year:
        year = start.year
        if start == startOf(year) and end == endOf(year):
            timePeriodStr = end.strftime('%Y')
        elif start.month == end.month:
            month = start.month
            if start == startOf(year, month) and end == endOf(year, month):
                timePeriodStr = end.strftime('%B %Y')
            elif start.day == end.day:
                timePeriodStr = dayAsStr(end) + end.strftime(' %B %Y')
            else:
                timePeriodStr = dayAsStr(start) + \
                        ' - ' + dayAsStr(end) + end.strftime(' %B %Y')
        else:
            if start == startOf(year, start.month) and end == endOf(year, end.month):
                timePeriodStr = start.strftime('%B') + ' - ' + end.strftime('%B %Y')
            else:
                timePeriodStr = dayAsStr(start) + start.strftime(' %B') + \
                        ' - ' + dayAsStr(end) + end.strftime(' %B %Y')
    else:
        if start == startOf(start.year) and end == endOf(end.year):
            timePeriodStr = start.strftime('%Y') + ' - ' + end.strftime('%Y')
        elif start == startOf(start.year, start.month) and end == endOf(end.year, end.month):
                timePeriodStr = start.strftime('%B %Y') + ' - ' + end.strftime('%B %Y')
        else:
            timePeriodStr = dayAsStr(start) + start.strftime(' %B %Y') + \
                    ' - ' + dayAsStr(end) + end.strftime(' %B %Y')

    return timePeriodStr

def size(bytes):
    if bytes < 1024:
        return '%d bytes' % bytes
    kiloBytes = float(bytes) / 1024
    if kiloBytes < 1024:
        return '%.2f KB' % kiloBytes
    megaBytes = kiloBytes / 1024
    if megaBytes < 1024:
        return '%.2f MB' % megaBytes
    gigaBytes = megaBytes / 1024
    if gigaBytes < 1024:
        return '%.2f GB' % gigaBytes

def getValue(record, key='pages'):
    return record[1][key]

def printerBreakdown(dict):
    if options.printerCounts and (options.userCounts or options.groupCounts):
        printerBreakdowns = ['Per printer:\n============']
        for name, stats in dict.dict.iteritems():
            users = stats.users
            groups = stats.groups
            printerBreakdowns.append('\n%s\n%s\n' % (name, '-' * len(name)))
            if options.userCounts:
                printerBreakdowns.append('\nUsers:\n~~~~~~\n::\n\n')
                printerBreakdowns.append(users.summarise())
            if options.groupCounts:
                printerBreakdowns.append('\nGroups:\n~~~~~~~\n::\n\n')
                printerBreakdowns.append(groups.summarise())
        return ''.join(printerBreakdowns)
    else:
        return ''

def makeTitle(subject):
    titleUnderline = '=' * (len(subject) + 2)
    return '%s\n %s\n%s' % (titleUnderline, subject, titleUnderline)

def createMail (mailFrom, mailTo, html, text, subject):
    from email.mime.text import MIMEText
    if html:
        from email.mime.multipart import MIMEMultipart

        message = MIMEMultipart('alternative')
        message['Subject'] = subject
        message['From'] = mailFrom
        message['To'] = ','.join(mailTo)

        message.attach(MIMEText(text, 'plain'))
        message.attach(MIMEText(html, 'html'))
    else:
        message = MIMEText(text, 'plain')
        message['Subject'] = subject
        message['From'] = mailFrom
        message['To'] = ','.join(mailTo)
    return message.as_string()

parser = OptionParser()
parser.add_option('-t', '--time-period', dest='timePeriod',
                  help='The time period to summarise (default: the previous full month).\n' \
                  'Format: YYYY[-MM[-DD]] or \'today\'.\n' \
                  'Examples: \'2009\' for a full year, \'2009-01-01\' for a single day',
                  default=None)
parser.add_option('-s', '--start-date', dest='startDate',
                  help='Start summary from the beginning of DATE.\n'
                  'Has no effect if --time-period is specified.\n' \
                  'Format: YYYY-MM-DD', metavar='DATE', default=None)
parser.add_option('-e', '--end-date', dest='endDate',
                  help='Finish summary at the end of DATE.\n'
                  'Has no effect if --time-period is specified.\n' \
                  '(default if --start-date is specified: %default).\n' \
                  'Format: YYYY-MM-DD', metavar='DATE', default='today')
parser.add_option('-p', '--print-server', dest='printServers', action='append',
                  help='Read print events from SERVER (may be specified multiple times; ' \
                  'default: \'localhost\')',
                  metavar='SERVER', default=['localhost'])
parser.add_option('-i', '--ignore-printer', dest='ignorePrinters', action='append',
                  help='Ignore jobs printed on PRINTER (may be specified multiple times)',
                  metavar='PRINTER', default=[])
parser.add_option('-m', '--mail-to', dest='mailTo', action='append',
                  help='Send report to ADDRESS (may be specified multiple times)',
                  metavar='ADDRESS', default=None)
parser.add_option('-d', '--details', help='Show full print log details (default: %default)',
                  action='store_true', default=False)
parser.add_option('--stdout', help='Print report to stdout ' \
                  '(default: False, unless no mail recipients are given)',
                  action='store_true', default=False)
parser.add_option('--mail-server', help='SMTP server to use for e-mail reports. ' \
                  '(default: %default)',
                  metavar='SERVER', dest='mailServerName', default='localhost')
parser.add_option('--mail-from', dest='mailFrom', default='Administrator',
                  help="'From' address to use for mailed reports.\n"
                  '(default: %default)', metavar='ADDRESS')
parser.add_option('--users', help='Generate per-user summaries.\n' \
                  '(default: %default)', dest='userCounts',
                  action='store_true', default=True)
parser.add_option('--no-users', dest='userCounts', action='store_false')
parser.add_option('--groups', help='Generate per-group summaries.\n' \
                  '(default: %default)', dest='groupCounts',
                  action='store_true', default=True)
parser.add_option('--no-groups', dest='groupCounts', action='store_false')
parser.add_option('--printers', help='Generate per-printer summaries.\n' \
                  '(default: %default)', dest='printerCounts',
                  action='store_true', default=True)
parser.add_option('--no-printers', dest='printerCounts', action='store_false')
parser.add_option('-r', '--raw', help='Output raw markup. (default: %default)',
                  dest='rawOutput', action='store_true', default=False)
options, args = parser.parse_args()
if options.groupCounts:
    try:
        import active_directory
    except ImportError:
        print('Couldn\'t load \'active_directory\' module; disabling group summaries.')
        options.groupCounts = False;


def generateReport(options):
    eventLogStartTime = getStartDate(options)
    eventLogEndTime = getEndDate(options)

    subject = 'Print statistics for %s' \
            % getTimePeriodAsStr(eventLogStartTime, eventLogEndTime)
    sections = [makeTitle(subject)]
    errors = []

    eventSource = 'Print'
    eventID = 10 # Print notification
    eventLogFlags = win32evtlog.EVENTLOG_BACKWARDS_READ | \
                    win32evtlog.EVENTLOG_SEQUENTIAL_READ
    eventLogType = 'System'
    eventRegEx = re.compile(
        r'Document (?P<number>\d+), ((?P<application>.+) - )?(?P<name>.+) owned ' \
        'by (?P<user>\w+) was printed on (?P<printer>.+) via port (?P<port>.+)\.' \
        ' +Size in bytes: (?P<bytes>\d+); pages printed: (?P<pages>\d+)'
    )

    if options.details:
        printList = ['Details:\n========\n::\n\n']

    totals = Record()
    userDict = JobDict('User')
    groupDict = JobDict('Group')
    printerDict = JobDict('Printer', PrinterRecord)

    for printServer in options.printServers:
        eventLogHandle = win32evtlog.OpenEventLog(printServer, eventLogType)
        eventsRemaining = True
        eventTime = eventLogEndTime
        while eventsRemaining:
            events = win32evtlog.ReadEventLog(eventLogHandle, eventLogFlags, 0)
            for event in events:
                eventTime = datetime.strptime(event.TimeGenerated.Format(),
                                              '%m/%d/%y %H:%M:%S')
                if eventTime > eventLogEndTime:
                    continue
                elif eventTime < eventLogStartTime:
                    eventsRemaining = False
                    break
                elif winerror.HRESULT_CODE(event.EventID) == eventID and \
                     event.SourceName == eventSource:
                    eventMessage = win32evtlogutil.SafeFormatMessage(event,
                                                                     eventLogType)
                    match = eventRegEx.match(eventMessage)
                    if match:
                        userName = match.group('user')
                        printerName = match.group('printer')
                        if printerName in options.ignorePrinters:
                            continue
                        bytes = int(match.group('bytes'))
                        pages = int(match.group('pages'))
                        totals.addJob(pages, bytes)
                        if options.printerCounts:
                            printerDict.addJob(printerName, pages, bytes)
                        if options.userCounts:
                            userDict.addJob(userName, pages, bytes)
                            if options.printerCounts:
                                printerDict.dict[printerName].users.addJob(userName, pages, bytes)
                        if options.groupCounts:
                            # FIXME: Need to decide what to do about users not in a group
                            ADUser = active_directory.find_user(userName) or None
                            if ADUser and ADUser.memberOf:
                                for ADGroup in ADUser.memberOf:
                                    groupName = ADGroup.cn
                                    groupDict.addJob(groupName, pages, bytes)
                                    if options.printerCounts:
                                        printerDict.dict[printerName].groups.addJob(groupName, pages, bytes)
                        if options.details:
                            number = int(match.group('number'))
                            port = match.group('port')
                            eventMessage = 'Document %d owned by %s was printed on %s ' \
                                    'via port %s. Size in bytes: %d; pages printed: %d\n' \
                                    % (number, userName, printerName, port, bytes, pages)
                            printList.append(' %s: %s' % (eventTime, eventMessage))
                    else:
                        errors.append('Error: could not parse event message\n\t%s' \
                            % eventMessage)
            if not events:
                eventsRemaining = False
                if eventTime > eventLogStartTime:
                    errors.append('No events found on %s prior to %s' %
                                  (printServer, dayAsStr(eventTime) +
                                  eventTime.strftime(' %B %Y')))

    if options.printServers:
        sections.append('::\n')
        sections.append(' Print servers queried: ' + ', '.join(options.printServers) + '\n')

    if len(errors):
        sections.append('::\n')
        sections.append(' ' + '\n '.join(errors) + '\n')

    if totals.jobs:
        sections.append('Totals:\n=======\n::\n')
        sections.append(' Page count: %d' % totals.pages)
        sections.append(' Job count: %d' % totals.jobs)
        sections.append(' Average pages per job: %.2f' %
                        (float(totals.pages) / totals.jobs))
        sections.append(' Total print size: ' + size(totals.bytes))
        sections.append(' Average print size: ' +
                        size(float(totals.bytes) / totals.jobs))
        sections.append(' Note: Duplexed print jobs are not recorded, so'
                        ' contribute their full page count as if single-sided.\n')

        if options.userCounts:
            sections.append(userDict.summarise(True))
        if options.groupCounts:
            sections.append(groupDict.summarise(True))
        if options.printerCounts:
            sections.append(printerDict.summarise(True))
        sections.append(printerBreakdown(printerDict))
        if options.details:
            sections.append(''.join(printList))

    return '\n'.join(sections), subject


mailBody, mailSubject = generateReport(options)


if options.rawOutput:
    mailText = mailBody
else:
    mailText = mailBody.replace('::\n', '')

if options.stdout or not options.mailTo:
    print mailText

if options.mailTo:
    try:
        from docutils.core import publish_string
        mailHtml = publish_string(source = mailBody, writer_name = 'html')
    except ImportError:
        print('Couldn\'t load \'docutils\' module; disabling HTML e-mail.')
        mailHtml = None
    mailMessage = createMail(options.mailFrom, options.mailTo, mailHtml, mailText, mailSubject)
    mailServer = smtplib.SMTP(options.mailServerName)
    mailServer.sendmail(options.mailFrom, options.mailTo, mailMessage)
    mailServer.quit()
