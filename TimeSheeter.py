from __future__ import print_function
import datetime
import re

import dateutil
import dateutil.parser
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

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

def days_hours_minutes(td):
    return td.days, td.seconds//3600, (td.seconds//60)%60

def duration_hours_minutes(td):
    days, hours, minutes = days_hours_minutes(td)
    hours = (days *24) + hours
    return hours, minutes

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
    client_short_long = ("@sma", "SmartQare")
    # date and time string formats:
    dayformat = "%d"
    timeformat = "%H:%M"
    datetimeformat = "%d-%m-%Y %H:%M"

    ###########################################

    print("Client: " + client_short_long[1])
    # Call the Calendar API
    #now = datetime.datetime.utcnow().isoformat() # 'Z' indicates UTC time
    now = datetime.datetime.utcnow()
    firstDaylastMonth = now.replace(day=1, hour=0, minute=1) - datetime.timedelta(days=1)
    lastDaylastMonth = firstDaylastMonth.replace(hour=23, minute=59)
    firstDaylastMonth = firstDaylastMonth.replace(day=1)

    # add Z
    now = now.isoformat() + 'Z'
    lastDaylastMonth = lastDaylastMonth.isoformat() + 'Z'
    firstDaylastMonth = firstDaylastMonth.isoformat() + 'Z'

    events_result = service.events().list(calendarId=CalID, timeMin=firstDaylastMonth, timeMax=lastDaylastMonth,
                                        maxResults=1000, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    total_duration = datetime.timedelta()
    print("Day" + "\t" + "Start time" + "\t" + "End time" + "\t" + "Duration (hours:minutes)" + "\t" + "Description" + "\t")

    if not events:
        print('No upcoming events found.')
    for event in events:
        #print header of table column names

        eventsummary = event['summary']
        if eventsummary.find(client_short_long[0]) > -1:
            #remove client abbreviation. Delete's any word that partially matches the string in client_short_long[0]
            eventsummary = re.sub(r"" + client_short_long[0] + "\w+", "", eventsummary).lstrip() #eventsummary.replace(client_short_long[0], "")
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
                print(start_parsed.strftime(dayformat) + "\t" + start_parsed.strftime(timeformat) + "\t" + end_parsed.strftime(timeformat) + "\t" + str(
                    duration_hours) + ":" + str(duration_minutes) + "\t" + eventsummary)
    hours, minutes = duration_hours_minutes(total_duration)

    print("Total duration was: " + str(hours) + " hours and " + str(minutes) + " minutes")



if __name__ == '__main__':
    main()