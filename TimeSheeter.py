from __future__ import print_function
import datetime
import re
import calendar
import dateutil
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
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
import configparser
import argparse
from dataclasses import dataclass, field

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

now = datetime.datetime.utcnow()
# last month
last_day_last_month = now.replace(day=1, hour=23, minute=59, second=59) - datetime.timedelta(days=1)
first_day_last_month = last_day_last_month.replace(day=1, hour=0, minute=0, second=0)
# this month
first_day_this_month = now.replace(day=1, hour=0, minute=0, second=0)
last_day_this_month = (first_day_this_month + relativedelta(months=1)) - datetime.timedelta(days=1)

# initiate a dataframe to hold timesheet dataclass
# this is used here kind of like a C++ typedef for Python
@dataclass(init=True)
class TimeSheetData:
    client_name : str
    total_duration : field(default_factory=lambda: datetime.timedelta())
    #dfcolumns = ["Day", "Start_time", "End_time", "Duration", "Week_nr", "Week_duration", "Description"]
    time_sheet_df : field(default_factory=lambda: pd.DataFrame())

def get_script_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

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

def print_datatable(time_table : pd.DataFrame):
    print(tabulate(time_table, headers=time_table.head(), tablefmt="presto"))

def is_date(string : str = "", fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try:
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False

# Default Time zone is Amsterdam 'GMT+01:00'
def get_gcal_events(start_date: datetime  = first_day_last_month, end_date : datetime = last_day_last_month, timeZone = 'GMT+01:00'):
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
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
    # TODO: add as CLI arguments.
    # TODO: Add capability for searching by calendar name.
    # Calendar ID. You can find this ID when you go to google calendar (each calendar has it's own ID)
    CalID = 'xaop.com_g07i0og0nf1tortohmakm4cmak@group.calendar.google.com'
    ###########################################

    # Call the Calendar API
    # Convert start and end dates to iso strings and add Z
    # example date string: "2019-11-30T23:59:55.783958Z"
    start_date = start_date.isoformat() + 'Z'
    end_date = end_date.isoformat() + 'Z'

    print("Getting gcal events between " + start_date + " and " + end_date)

    events_result = service.events().list(calendarId=CalID, timeMin=start_date, timeMax=end_date,
                                          maxResults=1000, singleEvents=True,
                                          orderBy='startTime',
                                          timeZone=timeZone).execute()
    events = events_result.get('items', [])

    return events

# searches for tagged words in list of words.
# example: returns Waverboard for tag_str = @wav and input_list = "bloe bla bloa Waverboard asdfds"
# because the wav tag is in Waverboard
# returns "" if none found
def get_tagword_from_list(input_list, tag_str):
    out_str = ""
    # remove the tag char (first character) from tag_str:
    tag_str = tag_str[1:]
    for in_str in input_list:
        if tag_str.lower() in in_str.lower():
            out_str = in_str
            break

    return out_str

# gets a tagged word from a string.
# Example: get_tag_in_str("blabla bla bla foo foo foo @kik bla bla", "@") would return @kik
# returns "" if none found
def get_tag_in_str(input_str, tag = "@"):
    out_str = ""
    word_list = input_str.split()
    for word in word_list:
        if word.lower().startswith(tag):
            out_str = word
            break

    return out_str

# Gets all tagged words from a stringlist
# Example: get_strlist_of_tags("this @is a string with some @words", "@")
# Returns ["@is", "@words"]
def get_strlist_of_tags(input_list, tag = "@"):
    out_list = []
    for input_str in input_list:
        word_list = input_str.split()
        for word in word_list:
            if word.lower().startswith(tag):
                if word not in out_list:
                    out_list.append(word)

    return out_list

# creates a dict of all long client names corresponding to tags in tag_list
# If there are multiple client names in client_list that match one or more tags, only the first client name is returned
# Example:
# tag_list = ['@wav', '@cli', '@clie', '@clien']
# client_list = Wavin, Test, Client name, client name, client_name, cli
# Returns {'@wav': 'Wavin', '@cli': 'cli', '@clie': 'Client name', '@clien': 'Client name'}
def get_client_tag_dict(tag_list, client_list):
    #out_lst = []
    out_dct = {}

    # order wordlist by word length
    client_list.sort(key=len)

    for tag in tag_list:
        for client in client_list:
            if tag[1:].lower() in client.lower():
                already_in_out = False
                #for out in out_dct:
                if out_dct.get(tag) != None:
                    already_in_out = True
                if not already_in_out:
                    out_dct.update({tag: client})

    return out_dct

# creates a dict of all long client names and their corresponding tags
# One client name can have multiple matching tags. Example: "Client name" can have @cli and @Clien as matching tags
def get_client_name_dict(tag_list, client_list):

    client_tag_dict = get_client_tag_dict(tag_list, client_list)

    # reverse key, values in client_tag_dict
    inv_map = {v: [i for i in client_tag_dict.keys() if client_tag_dict[i] == v] for k, v in client_tag_dict.items()}

    return inv_map


def get_timesheet(start_date: str = first_day_last_month, end_date : str = last_day_last_month, client_list: list = [], week_totals : bool = False, strFreelancer = "Georges Meinders"):

    # date and time string formats:
    # TODO: make configurable in ini file?
    dayformat = "%d"
    timeformat = "%H:%M"
    daymonthformat = "%d-%m"
    monthyearformat = "%m - %Y"  # for month indication on timesheet
    yearmonthformat = "%Y %m"  # for timesheet filename
    datetimeformat = "%d-%m-%Y %H:%M"
    timeZone = 'GMT+01:00'
    total_duration = datetime.timedelta()
    # TODO: make column names customizable through ini config file
    dfcolumns = ["Day", "Start_time", "End_time", "Duration", "Week_nr", "Week_duration", "Description"]
    #time_table = pd.DataFrame(columns=dfcolumns)
    time_table = TimeSheetData("", datetime.timedelta(), pd.DataFrame(columns=dfcolumns))
    time_tables = []

    # check input
    valid_start_d = is_date(str(start_date), True)
    valid_end_d = is_date(str(end_date), True)
    if not valid_start_d and valid_end_d:
        if not valid_start_d:
            print("Invalid start date entered:   " + start_date)
        if not valid_end_d:
            print("Invalid end date entered:   " + end_date)
        print("Quitting...")
        quit()

    events = get_gcal_events(start_date, end_date, timeZone=timeZone)

    # get all "@" tagged words from all event descriptions
    # First create list of all event description strings for the entire period
    event_summary_list = []
    for event in events:
        event_summary_list.append(event['summary'])
    # Get a list of all @tag tags in all events
    tag_list = []
    tag_list = get_strlist_of_tags(event_summary_list, "@")
    # Get a list of all client names corresponding to the @ tags
    client_tag_name_dct = get_client_tag_dict(tag_list, client_list)
    client_name_tag_dct = get_client_name_dict(tag_list, client_list)
    print("The following @tag client name matches where made: " + str(client_name_tag_dct))
    print("If you don't want a certain tag/client to be included, remove the client from client.ini file or add # in front of the client name.")

    if not events:
        print("No events found between " + start_date + " and " + end_date)
    else:
        # iterate through all events for each tag found (generating 1 report per tag/client)
        for cl_name in client_name_tag_dct:
            #previous_day = 0
            previous_week = 0
            week_dur = datetime.timedelta()  # reset to 0 duration
            for event in events:
                eventsummary = event['summary']
                # for each @tag match for the current client name iteration:
                for cl_tag in client_name_tag_dct.get(cl_name):
                    # if the client tag is found in the calendar event description
                    if cl_tag.lower() in eventsummary.lower():
                        # Add client name to the timesheet object
                        if len(time_table.time_sheet_df) == 0:
                            # add client name to TimeSheetData object
                            time_table.client_name = cl_name
                            print("Generating time sheet for client: " + cl_name)

                        # To remove @clientname tag from event summary
                        eventsummary = re.sub(r"" + cl_tag + "\W+", "", eventsummary,
                                              flags=re.I).lstrip()
                        eventsummary = re.sub(r"" + cl_tag + "\w+", "", eventsummary, flags=re.I).lstrip()

                        # time handling stuff:
                        # get datetime string of start of event
                        start = event['start'].get('dateTime', event['start'].get('date'))
                        # get datetime string of end of event
                        end = event['end'].get('dateTime', event['end'].get('date'))
                        # convert time strings to dateutil types
                        start_parsed = dateutil.parser.parse(start)
                        end_parsed = dateutil.parser.parse(end)
                        # calculate duration of each event
                        duration_event = end_parsed - start_parsed
                        # convert duration to total hours and minutes (standard str convert converts to days hours and minutes,
                        # we want total time in hours instead of days
                        duration_hours, duration_minutes = duration_hours_minutes(duration_event)
                        # add up hours of this event to all hours of all events
                        total_duration = total_duration + duration_event
                        # if event spans multiple days:
                        # TODO: this bit for events that span multiple days is probably wrong now.
                        # should go into dataframe, not tab seperated.
                        if start_parsed.day > end_parsed.day:
                            print(start_parsed.strftime(dayformat) + " - " + end_parsed.strftime(
                                dayformat) + "\t" + start_parsed.strftime(
                                timeformat) + "\t" + end_parsed.strftime(timeformat) + "\t" + str(
                                duration_hours) + ":" + str(duration_minutes) + "\t" + eventsummary)
                        # if event within the same day:
                        else:
                            # ["Day", "Start_time", "End_time", "Duration", "Week_nr", "Week_duration", "Description"]

                            # Set day string
                            day = start_parsed.strftime(dayformat)

                            # Set week nr string
                            week_nr = start_parsed.strftime("%V")
                            # add up total hours for this week so far:
                            if not (int(week_nr) == previous_week):
                                week_dur = datetime.timedelta()  # reset to 0 duration
                                previous_week = int(week_nr)
                            week_dur = week_dur + datetime.timedelta(hours=duration_hours, minutes=duration_minutes)
                            hours, minutes = duration_hours_minutes(week_dur)
                            week_dur_str = str(hours) + " : " + str(minutes)

                            row = pd.DataFrame({'Day': day,
                                                'Start_time': [start_parsed.strftime(timeformat)],
                                                'End_time': [end_parsed.strftime(timeformat)],
                                                'Duration': [(str(duration_hours) + ":" + str(duration_minutes))],
                                                'Week_nr': week_nr,
                                                'Week_duration': week_dur_str,
                                                'Description': [eventsummary]})
                            # row = {start_parsed.strftime(dayformat), start_parsed.strftime(timeformat), end_parsed.strftime(timeformat), (str(duration_hours) + ":" + str(duration_minutes)),  eventsummary}
                            time_table.time_sheet_df = time_table.time_sheet_df.append(row, ignore_index=True)

            time_table.total_duration = total_duration
            time_tables.append(time_table)
            #clear/drop/reset TimeSheetData object to make it ready to add data of next client
            time_table = TimeSheetData("", datetime.timedelta(), pd.DataFrame(columns=dfcolumns))
            total_duration = datetime.timedelta()

    return time_tables

def parseargs():
    # Construct an argument parser
    all_args = argparse.ArgumentParser()

    # Add arguments to the parser
    #start date
    all_args.add_argument("-s", "--start", required=False, default="",
                          help="Start date. Script will assume time of day 0:00:00 unless otherwise specified.")
    #end date
    all_args.add_argument("-e", "--end", required=False, default="",
                          help="End date. Script will assume time of day 23:59:59 unless otherwise specified.")
    # use this month?
    all_args.add_argument("-l", "--last", action='store_true', default=False,
                          help="Generate time sheet for all of last month.")
    # use last month?
    all_args.add_argument("-t", "--this", action='store_true', default=False,
                          help="Generate time sheet for all of this month.")
    # generate week totals?
    all_args.add_argument("-w", "--weektotals", action='store_true', default=False,
                          help="Add week numbers and totals per week to report.")
    # generate report?
    all_args.add_argument("-r", "--report", action='store_true', default=False,
                          help="Generate a .pdf report based on the .docx template file in this directory.")

    # parse arguments
    args = vars(all_args.parse_args())

    # get start and end date from parsed arguments:
    if (is_date(args['start']) and is_date(args['end'])):
        start_date = args['start']
        end_date = args['end']
    elif args['this']:
        start_date = first_day_this_month
        end_date = last_day_this_month
    elif args['last']:
        start_date = first_day_last_month
        end_date = last_day_last_month
    else: #default is to generate report for last month
        start_date = first_day_last_month
        end_date = last_day_last_month

    report_week_totals = args['weektotals']
    make_report = args['report']

    return start_date, end_date, report_week_totals, make_report

# get a list of client names from a clients.ini
def get_clients():
    #script_path = get_script_path()

    config = configparser.RawConfigParser(allow_no_value=True)
    # keep keys case sensitive:
    config.optionxform = lambda option: option
    try:
        # parse file
        config.read('clients.ini')
    except Exception as e:
        print("\nError while trying to read client list.")
        print("Make sure a client.ini file exists with contents in the following format:")
        print("[client list]\nClient name 1\nClient name 2\nClient name 3\n\n")
        print(e)
        print("quitting")
        quit()

    client_list = list(config['client list'].keys())
    return client_list

def main():
    """Shows basic usage of the Google Calendar API.

    """
    start_date, end_date, report_week_totals, make_report = parseargs()
    client_list = get_clients()
    # get list of timesheets
    timesheets = get_timesheet(start_date, end_date, client_list, report_week_totals)
    for sheet in timesheets:
        print("\nTime sheet for client: " + sheet.client_name)
        hours, minutes = duration_hours_minutes(sheet.total_duration)
        print("Total duration for cient was: " + str(hours) + " hours and " + str(minutes) + " minutes.")
        print(tabulate(sheet.time_sheet_df, headers=sheet.time_sheet_df.head(), tablefmt="presto"))

    quit()

    ####################################################
    # After this is old code only.
    # TODO: dump pdf generation lines into a new function and implement argparse for it.





    # creds = None
    # # The file token.pickle stores the user's access and refresh tokens, and is
    # # created automatically when the authorization flow completes for the first
    # # time.
    # if os.path.exists('token.pickle'):
    #     with open('token.pickle', 'rb') as token:
    #         creds = pickle.load(token)
    # # If there are no (valid) credentials available, let the user log in.
    # if not creds or not creds.valid:
    #     if creds and creds.expired and creds.refresh_token:
    #         creds.refresh(Request())
    #     else:
    #         # TODO: nicer way for user to generate credentials? Login with user/pass?
    #         # Having to download credentials file from Google dev console sucks.
    #         flow = InstalledAppFlow.from_client_secrets_file(
    #             'client_secrets.json', SCOPES)
    #         creds = flow.run_local_server(port=0)
    #     # Save the credentials for the next run
    #     with open('token.pickle', 'wb') as token:
    #         pickle.dump(creds, token)
    #
    # service = build('calendar', 'v3', credentials=creds)
    #
    # ###################### Enter info:
    # # Calendar ID. You can find this ID when you go to google calendar (each calendar has it's own ID)
    # CalID = 'xaop.com_g07i0og0nf1tortohmakm4cmak@group.calendar.google.com'
    # #search string (@tag) + what the tag stands for
    # #client_short_long = ("@sma", "xxxxx")
    # client_short_long = ("@wav", "xxxx")
    # strFreelancer = "Georges Meinders"
    # # date and time string formats:
    # dayformat = "%d"
    # timeformat = "%H:%M"
    # monthyearformat = "%m - %Y" #for month indication on timesheet
    # yearmonthformat = "%Y %m" #for timesheet filename
    # datetimeformat = "%d-%m-%Y %H:%M"
    # # Time zone is Amsterdam
    # timeZone = 'GMT+01:00'
    #
    # ###########################################
    #
    # print("Client: " + client_short_long[1])
    # # Call the Calendar API
    # #now = datetime.datetime.utcnow().isoformat() # 'Z' indicates UTC time
    # now = datetime.datetime.utcnow()
    # #last month
    # firstDaylastMonth = now.replace(day=1, hour=0, minute=1) - datetime.timedelta(days=1)
    # lastDaylastMonth = firstDaylastMonth.replace(hour=23, minute=59)
    # firstDaylastMonth = firstDaylastMonth.replace(day=1)
    #
    # # other date placeholder
    # # this month:
    # firstDaylastMonth = now.replace(day=1, hour=0, minute=1) #actually first day this month
    # lastDayDateThisMonth = calendar.monthrange(now.year, now.month)[1]
    # lastDaylastMonth = now.replace(day=lastDayDateThisMonth, hour=23, minute=59) #actually last day this month
    #
    # # Set month str for document:
    # # Format = MM - yyyy
    # #strMonth = firstDaylastMonth.month
    #
    # # add Z
    # now = now.isoformat() + 'Z'
    # lastDaylastMonth = lastDaylastMonth.isoformat() + 'Z'
    # firstDaylastMonth = firstDaylastMonth.isoformat() + 'Z'
    #
    # # custom date:
    # #lastDaylastMonth = "2019-11-30T23:59:55.783958Z"
    # #firstDaylastMonth = "2019-11-23T23:59:55.783958Z"
    #
    # # added timeZone
    # events_result = service.events().list(calendarId=CalID, timeMin=firstDaylastMonth, timeMax=lastDaylastMonth,
    #                                     maxResults=1000, singleEvents=True,
    #                                     orderBy='startTime',
    #                                     timeZone=timeZone).execute()
    # events = events_result.get('items', [])
    #
    # total_duration = datetime.timedelta()
    #
    # #print("Day" + "\t" + "Start_time" + "\t" + "End_time" + "\t" + "Duration" + "\t" + "Description" + "\t")
    #
    # #TODO: convert to dataframe
    # dfcolumns = ["Day", "Start_time", "End_time", "Duration", "Description"]
    # time_table = pd.DataFrame(columns=dfcolumns)
    #
    # if not events:
    #     print('No upcoming events found.')
    # else:
    #     for event in events:
    #         #print header of table column names
    #
    #         eventsummary = event['summary']
    #         #if eventsummary.find(client_short_long[0]) > -1:
    #         if re.search(client_short_long[0], eventsummary, re.IGNORECASE):  #removed case sensitivity
    #             #remove client abbreviation. Delete's any word that partially matches the string in client_short_long[0]
    #             #added flags=re.I (for case insensitive match) and \W
    #             # Do this in 2 passes with \W and \w, because I do not get how regex works :P
    #             eventsummary = re.sub(r"" + client_short_long[0] + "\W+", "", eventsummary, flags=re.I).lstrip() #eventsummary.replace(client_short_long[0], "")
    #             eventsummary = re.sub(r"" + client_short_long[0] + "\w+", "", eventsummary, flags=re.I).lstrip()
    #             # get datetime string of start of event
    #             start = event['start'].get('dateTime', event['start'].get('date'))
    #             # get datetime string of end of event
    #             end = event['end'].get('dateTime', event['end'].get('date'))
    #             # convert time strings to dateutil types
    #             start_parsed = dateutil.parser.parse(start)
    #             end_parsed = dateutil.parser.parse(end)
    #             # calculate duration of each event
    #             duration = end_parsed - start_parsed
    #             # convert duration to total hours and minutes (standard str convert converts to days hours and minutes,
    #             # we want total time in hours instead of days
    #             duration_hours, duration_minutes = duration_hours_minutes(duration)
    #             # add up hours of this event to all hours of all events
    #             total_duration = total_duration + duration
    #             # if event spans multiple days:
    #             if start_parsed.day > end_parsed.day:
    #                 print(start_parsed.strftime(dayformat) + " - " + end_parsed.strftime(dayformat) + "\t" + start_parsed.strftime(
    #                     timeformat) + "\t" + end_parsed.strftime(timeformat) + "\t" + str(
    #                     duration_hours) + ":" + str(duration_minutes) + "\t" + eventsummary)
    #             # if event within the same day:
    #             else:
    #                 #row = pd.Series([start_parsed.strftime(dayformat), start_parsed.strftime(timeformat), end_parsed.strftime(timeformat), (str(duration_hours) + ":" + str(duration_minutes)),  eventsummary]) #, index = time_table.columns
    #                 # ["Day", "Start_time", "End_time", "Duration", "Description"]
    #                 row = pd.DataFrame({'Day' : [start_parsed.strftime(dayformat)],
    #                                     'Start_time': [start_parsed.strftime(timeformat)],
    #                                     'End_time': [end_parsed.strftime(timeformat)],
    #                                     'Duration':[(str(duration_hours) + ":" + str(duration_minutes))],
    #                                     'Description':[eventsummary]})
    #                 #row = {start_parsed.strftime(dayformat), start_parsed.strftime(timeformat), end_parsed.strftime(timeformat), (str(duration_hours) + ":" + str(duration_minutes)),  eventsummary}
    #                 time_table = time_table.append(row, ignore_index=True)
    #                 #time_table = pd.concat(time_table, row, ignore_index=True)
    #                 #print(start_parsed.strftime(dayformat) + "\t" + start_parsed.strftime(timeformat) + "\t" + end_parsed.strftime(timeformat) + "\t" + str(
    #                 #    duration_hours) + ":" + str(duration_minutes) + "\t" + eventsummary)
    #
    #     #print(time_table.to_string())
    #     print(tabulate(time_table, headers=time_table.head(), tablefmt="presto"))
    #
    #     hours, minutes = duration_hours_minutes(total_duration)
    #     print("Total duration was: " + str(hours) + " hours and " + str(minutes) + " minutes")
    #
    #     scriptpath = os.path.dirname(sys.argv[0])
    #     templatefilename = "template.docx"
    #     strTempDocxFilenameOut = "temp_docx.docx" #cannot contain spaces. Me n00b not understand.
    #     strTempPDFfilenameOut = "temp_docx.pdf"
    #     pdfresultfilename = start_parsed.strftime(yearmonthformat) + " uren Georges Meinders " + client_short_long[1] + ".pdf"
    #     templapath = scriptpath + "/" + templatefilename
    #     strTempDocxpath = scriptpath + "/" + strTempDocxFilenameOut
    #     strTempPDFpath = scriptpath + "/" + strTempPDFfilenameOut
    #     outputpath = "C:/Users/bever/Desktop/Sync/Business/MEGAHARD/Administratie/Digitale facturen en bonnen/2020/" + pdfresultfilename
    #
    #     #print(templapath)
    #     doc = MailMerge(templapath)
    #
    #     #Paste data about timesheet into Word template:
    #     #doc.get_merge_fields()
    #     doc.merge(
    #         ClientName=client_short_long[1],
    #         FreelancerName=strFreelancer,
    #         MMyyyy=start_parsed.strftime(monthyearformat),
    #         TotalHours=str(hours),
    #         TotalMinutes=str(minutes),
    #         TimeZone=timeZone)
    #
    #     #Paste timesheet dataframe into Word doc template:
    #     #Convert dataframe to a list of dictionaries (each dictionary in the list is a row)
    #     doc.merge_rows('Day', time_table.T.to_dict().values())
    #     ###
    #     doc.write(strTempDocxFilenameOut)
    #     convert_docx_to_pdf(strTempDocxpath)
    #     # remove temporary docx file
    #     os.remove(strTempDocxpath)
    #     # rename and move PDF file
    #     os.rename(strTempPDFpath, outputpath)


if __name__ == '__main__':
    main()