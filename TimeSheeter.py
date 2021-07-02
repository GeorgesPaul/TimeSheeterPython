from __future__ import print_function
import datetime
import re
import calendar
import dateutil
import dateutil.parser
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from tabulate import tabulate
import win32com.client as client
#from docxtpl import DocxTemplate
from mailmerge import MailMerge
import sys, os
import locale

from tempfile import NamedTemporaryFile

# About:
# This script downloads Google calendar events, finds all events with a "@something" tag
# and generates a time sheet table with
# total duration, day, start and end times, event duration and description.
# Currently the table is tab seperated so it can be pasted into Excel

#####TODO:
# better auth flow without requiring dev console stuff
# break up into nice functions. Load calendar events into Python object with proper types and string equivalents?
# accept arguments when running python file e.g. "Python TimeSheeter.py -getLastMonth -clientTag -clientName"
# add option to output straight to HTML, Word or Excel (with markup?)
#

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def convert_docx_to_pdf(inputpath:str):
    """Save a pdf of a docx file."""
    try:
        word = client.DispatchEx("Word.Application")
        target_path = inputpath.replace(".docx", r".pdf")

        word_doc = word.Documents.Open(inputpath)
        word_doc.SaveAs(target_path, FileFormat=17)
        word_doc.Close()
    except Exception as e:
            raise e
            word.Quit()
    finally:
            word.Quit()

def days_hours_minutes(td):
    return td.days, td.seconds//3600, (td.seconds//60)%60


def duration_hours_minutes(td):
    days, hours, minutes = days_hours_minutes(td)
    hours = (days *24) + hours
    return hours, minutes


def add_months(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year,month)[1])
    return datetime.date(year, month, day)

def main():
    """Shows basic usage of the Google Calendar API.

    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # TODO: nicer way for user to generate credentials? Login with user/pass?
            # Having to download credentials file from Google dev console sucks.
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secrets.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    ###################### Enter info:
    # Calendar ID. You can find this ID when you go to google calendar (each calendar has it's own ID)
    CalID = 'xaop.com_g07i0og0nf1tortohmakm4cmak@group.calendar.google.com'
    #search string (@tag) + what the tag stands for
    #client_short_long = ("@sma", "SmartQare")
    client_short_long = ("@wav", "Wavin")
    strFreelancer = "Georges Meinders"
    # date and time string formats:
    dayformat = "%d"
    timeformat = "%H:%M"
    monthyearformat = "%m - %Y" #for month indication on timesheet
    yearmonthformat = "%Y %m" #for timesheet filename
    datetimeformat = "%d-%m-%Y %H:%M"
    # Time zone is Amsterdam
    timeZone = 'GMT+01:00'

    ###########################################

    print("Client: " + client_short_long[1])
    # Call the Calendar API
    #now = datetime.datetime.utcnow().isoformat() # 'Z' indicates UTC time
    now = datetime.datetime.utcnow()
    #last month
    firstDaylastMonth = now.replace(day=1, hour=0, minute=1) - datetime.timedelta(days=1)
    lastDaylastMonth = firstDaylastMonth.replace(hour=23, minute=59)
    firstDaylastMonth = firstDaylastMonth.replace(day=1)

    # other date placeholder
    # this month:
    firstDaylastMonth = now.replace(day=1, hour=0, minute=1) #actually first day this month
    lastDayDateThisMonth = calendar.monthrange(now.year, now.month)[1]
    lastDaylastMonth = now.replace(day=lastDayDateThisMonth, hour=23, minute=59) #actually last day this month

    # Set month str for document:
    # Format = MM - yyyy
    #strMonth = firstDaylastMonth.month

    # add Z
    now = now.isoformat() + 'Z'
    lastDaylastMonth = lastDaylastMonth.isoformat() + 'Z'
    firstDaylastMonth = firstDaylastMonth.isoformat() + 'Z'


    # custom date:
    #lastDaylastMonth = "2019-11-30T23:59:55.783958Z"
    #firstDaylastMonth = "2019-11-23T23:59:55.783958Z"

    # added timeZone
    events_result = service.events().list(calendarId=CalID, timeMin=firstDaylastMonth, timeMax=lastDaylastMonth,
                                        maxResults=1000, singleEvents=True,
                                        orderBy='startTime',
                                        timeZone=timeZone).execute()
    events = events_result.get('items', [])

    total_duration = datetime.timedelta()

    #print("Day" + "\t" + "Start_time" + "\t" + "End_time" + "\t" + "Duration" + "\t" + "Description" + "\t")

    #TODO: convert to dataframe
    dfcolumns = ["Day", "Start_time", "End_time", "Duration", "Description"]
    time_table = pd.DataFrame(columns=dfcolumns)

    if not events:
        print('No upcoming events found.')
    else:
        for event in events:
            #print header of table column names

            eventsummary = event['summary']
            #if eventsummary.find(client_short_long[0]) > -1:
            if re.search(client_short_long[0], eventsummary, re.IGNORECASE):  #removed case sensitivity
                #remove client abbreviation. Delete's any word that partially matches the string in client_short_long[0]
                #added flags=re.I (for case insensitive match) and \W
                # Do this in 2 passes with \W and \w, because I do not get how regex works :P
                eventsummary = re.sub(r"" + client_short_long[0] + "\W+", "", eventsummary, flags=re.I).lstrip() #eventsummary.replace(client_short_long[0], "")
                eventsummary = re.sub(r"" + client_short_long[0] + "\w+", "", eventsummary, flags=re.I).lstrip()
                # get datetime string of start of event
                start = event['start'].get('dateTime', event['start'].get('date'))
                # get datetime string of end of event
                end = event['end'].get('dateTime', event['end'].get('date'))
                # convert time strings to dateutil types
                start_parsed = dateutil.parser.parse(start)
                end_parsed = dateutil.parser.parse(end)
                # calculate duration of each event
                duration = end_parsed - start_parsed
                # convert duration to total hours and minutes (standard str convert converts to days hours and minutes,
                # we want total time in hours instead of days
                duration_hours, duration_minutes = duration_hours_minutes(duration)
                # add up hours of this event to all hours of all events
                total_duration = total_duration + duration
                # if event spans multiple days:
                if start_parsed.day > end_parsed.day:
                    print(start_parsed.strftime(dayformat) + " - " + end_parsed.strftime(dayformat) + "\t" + start_parsed.strftime(
                        timeformat) + "\t" + end_parsed.strftime(timeformat) + "\t" + str(
                        duration_hours) + ":" + str(duration_minutes) + "\t" + eventsummary)
                # if event within the same day:
                else:
                    #row = pd.Series([start_parsed.strftime(dayformat), start_parsed.strftime(timeformat), end_parsed.strftime(timeformat), (str(duration_hours) + ":" + str(duration_minutes)),  eventsummary]) #, index = time_table.columns
                    # ["Day", "Start_time", "End_time", "Duration", "Description"]
                    row = pd.DataFrame({'Day' : [start_parsed.strftime(dayformat)],
                                        'Start_time': [start_parsed.strftime(timeformat)],
                                        'End_time': [end_parsed.strftime(timeformat)],
                                        'Duration':[(str(duration_hours) + ":" + str(duration_minutes))],
                                        'Description':[eventsummary]})
                    #row = {start_parsed.strftime(dayformat), start_parsed.strftime(timeformat), end_parsed.strftime(timeformat), (str(duration_hours) + ":" + str(duration_minutes)),  eventsummary}
                    time_table = time_table.append(row, ignore_index=True)
                    #time_table = pd.concat(time_table, row, ignore_index=True)
                    #print(start_parsed.strftime(dayformat) + "\t" + start_parsed.strftime(timeformat) + "\t" + end_parsed.strftime(timeformat) + "\t" + str(
                    #    duration_hours) + ":" + str(duration_minutes) + "\t" + eventsummary)

        #print(time_table.to_string())
        print(tabulate(time_table, headers=time_table.head(), tablefmt="presto"))

        hours, minutes = duration_hours_minutes(total_duration)
        print("Total duration was: " + str(hours) + " hours and " + str(minutes) + " minutes")

        scriptpath = os.path.dirname(sys.argv[0])
        templatefilename = "template.docx"
        strTempDocxFilenameOut = "temp_docx.docx" #cannot contain spaces. Me n00b not understand.
        strTempPDFfilenameOut = "temp_docx.pdf"
        pdfresultfilename = start_parsed.strftime(yearmonthformat) + " uren Georges Meinders " + client_short_long[1] + ".pdf"
        templapath = scriptpath + "/" + templatefilename
        strTempDocxpath = scriptpath + "/" + strTempDocxFilenameOut
        strTempPDFpath = scriptpath + "/" + strTempPDFfilenameOut
        outputpath = "C:/Users/bever/Desktop/Sync/Business/MEGAHARD/Administratie/Digitale facturen en bonnen/2020/" + pdfresultfilename

        #print(templapath)
        doc = MailMerge(templapath)

        #Paste data about timesheet into Word template:
        #doc.get_merge_fields()
        doc.merge(
            ClientName=client_short_long[1],
            FreelancerName=strFreelancer,
            MMyyyy=start_parsed.strftime(monthyearformat),
            TotalHours=str(hours),
            TotalMinutes=str(minutes),
            TimeZone=timeZone)

        #Paste timesheet dataframe into Word doc template:
        #Convert dataframe to a list of dictionaries (each dictionary in the list is a row)
        doc.merge_rows('Day', time_table.T.to_dict().values())
        ###
        doc.write(strTempDocxFilenameOut)
        convert_docx_to_pdf(strTempDocxpath)
        # remove temporary docx file
        os.remove(strTempDocxpath)
        # rename and move PDF file
        os.rename(strTempPDFpath, outputpath)


if __name__ == '__main__':
    main()