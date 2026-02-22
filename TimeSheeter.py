import datetime
import re
import calendar
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from tabulate import tabulate
import configparser
import argparse
from dataclasses import dataclass, field
import logging
from enum import Enum, auto
import yaml

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


@dataclass
class TimeSheetData:
    client_name: str
    total_duration: datetime.timedelta = field(default_factory=lambda: datetime.timedelta())
    time_sheet_df: pd.DataFrame = field(default_factory=lambda: pd.DataFrame())


class OutputFormat(Enum):
    TABLE = "table"
    CSV = "csv"
    TOTAL = "total"

    def __str__(self):
        return self.value


class TimesheetGenerator:
    def __init__(self):
        self.now = datetime.datetime.utcnow()
        self.last_day_last_month = self.now.replace(day=1, hour=23, minute=59, second=59) - datetime.timedelta(days=1)
        self.first_day_last_month = self.last_day_last_month.replace(day=1, hour=0, minute=0, second=0)
        self.first_day_this_month = self.now.replace(day=1, hour=0, minute=0, second=0)
        self.last_day_this_month = (self.first_day_this_month + relativedelta(months=1)) - datetime.timedelta(days=1)

        self.config = self.load_config()
        self.yaml_data = self.load_yaml()
        self.client_list = list(self.yaml_data['Clients'].keys())

    def load_config(self):
        config = configparser.ConfigParser()
        config.read('config.ini')
        return config
    
    def load_yaml(self):
        try:
            with open('clients.yaml', 'r') as file:
                return yaml.safe_load(file)
        except Exception as e:
            logging.error("Error reading clients.yaml: %s", str(e))
            raise

    def list_calendars(self):
        creds = self.get_credentials()
        service = build('calendar', 'v3', credentials=creds)

        calendar_list = service.calendarList().list().execute()
        calendars = calendar_list.get('items', [])

        if not calendars:
            print('No calendars found.')
        else:
            print('Calendars:')
            for calendar in calendars:
                print(f"ID: {calendar['id']}")
                print(f"Name: {calendar['summary']}")
                print(f"Description: {calendar.get('description', 'No description')}")
                print('-' * 40)
    def get_clients(self):
        config = configparser.RawConfigParser(allow_no_value=True)
        config.optionxform = lambda option: option
        try:
            config.read('clients.ini')
            return list(config['client list'].keys())
        except Exception as e:
            logging.error("Error reading client list: %s", str(e))
            print("Make sure a client.ini file exists with contents in the following format:")
            print("[client list]\nClient name 1\nClient name 2\nClient name 3\n")
            raise

    def get_credentials(self):
        from google.auth.exceptions import RefreshError
        creds = None
        if os.path.exists('token.pickle'):
            try:
                with open('token.pickle', 'rb') as token:
                    creds = pickle.load(token)
            except Exception as e:
                logging.warning(f"Failed to load token.pickle: {e}")
                creds = None
                
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except RefreshError:
                    logging.warning("Refresh token expired or invalid. Restarting auth process.")
                    flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        return creds

    def get_gcal_events(self, start_date, end_date, time_zone='GMT+01:00'):
        creds = self.get_credentials()
        service = build('calendar', 'v3', credentials=creds)

        cal_id = self.config.get('Google Calendar', 'CalID', fallback=None)
        if not cal_id:
            raise ValueError("CalID must be provided in the config.ini file")

        start_date = start_date.isoformat() + 'Z'
        end_date = end_date.isoformat() + 'Z'

        logging.info("Getting gcal events between %s and %s", start_date, end_date)

        events_result = service.events().list(calendarId=cal_id, timeMin=start_date, timeMax=end_date,
                                              maxResults=1000, singleEvents=True,
                                              orderBy='startTime', timeZone=time_zone).execute()
        return events_result.get('items', [])

    @staticmethod
    def get_strlist_of_tags(input_list, tag="@"):
        out_list = []
        for input_str in input_list:
            out_list.extend([word for word in input_str.split() if word.lower().startswith(tag)])
        return list(set(out_list))

    def get_client_tag_dict(self, tag_list, client_list):
        client_list.sort(key=len)
        out_dct = {}
        for tag in tag_list:
            for client in client_list:
                alias = self.yaml_data['Clients'].get(client, {}).get('alias_tag', client)
                if tag[1:].lower() in alias.lower():
                    if tag not in out_dct:
                        out_dct[tag] = client
                    break
        return out_dct

    def get_client_name_dict(self, tag_list, client_list):
        client_tag_dict = self.get_client_tag_dict(tag_list, client_list)
        return {v: [k for k, val in client_tag_dict.items() if val == v] for v in client_tag_dict.values()}

    def process_events(self, events, client_list):
        event_summary_list = [event['summary'] for event in events]
        tag_list = self.get_strlist_of_tags(event_summary_list, "@")
        client_name_tag_dict = self.get_client_name_dict(tag_list, client_list)

        if self.output_format == OutputFormat.TABLE:
            logging.info("The following @tag client name matches were made: %s", str(client_name_tag_dict))
            print(
                "If you don't want a certain tag/client to be included, remove the client from client.ini file or add # in front of the client name.")

        time_tables = []
        for cl_name, cl_tags in client_name_tag_dict.items():
            time_table = self.process_client_events(events, cl_name, cl_tags)
            if not time_table.time_sheet_df.empty:
                time_tables.append(time_table)

        return time_tables

    def process_client_events(self, events, client_name, client_tags):
        time_table = TimeSheetData(client_name)
        previous_week = 0
        week_dur = datetime.timedelta()

        for event in events:
            event_summary = event['summary']
            if any(tag.lower() in event_summary.lower() for tag in client_tags):
                if time_table.time_sheet_df.empty and self.output_format == OutputFormat.TABLE:
                    logging.info("Generating time sheet for client: %s", client_name)

                event_summary = self.clean_event_summary(event_summary, client_tags)
                start_parsed, end_parsed, duration_event = self.parse_event_times(event)

                week_nr, week_dur = self.update_week_duration(start_parsed, duration_event, previous_week, week_dur)
                previous_week = int(week_nr)

                row = self.create_event_row(start_parsed, end_parsed, duration_event, week_nr, week_dur, event_summary)
                #time_table.time_sheet_df = time_table.time_sheet_df.append(row, ignore_index=True) deprecated
                time_table.time_sheet_df = pd.concat([time_table.time_sheet_df, row], ignore_index=True)
                time_table.total_duration += duration_event

        return time_table

    @staticmethod
    def clean_event_summary(summary, tags):
        for tag in tags:
            summary = re.sub(rf"{tag}\W+", "", summary, flags=re.I).strip()
            summary = re.sub(rf"{tag}\w+", "", summary, flags=re.I).strip()
        return summary

    @staticmethod
    def parse_event_times(event):
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))
        start_parsed = parse(start)
        end_parsed = parse(end)
        duration_event = end_parsed - start_parsed
        return start_parsed, end_parsed, duration_event

    @staticmethod
    def update_week_duration(start_parsed, duration_event, previous_week, week_dur):
        week_nr = start_parsed.strftime("%V")
        if int(week_nr) != previous_week:
            week_dur = datetime.timedelta()
        week_dur += duration_event
        return week_nr, week_dur

    @staticmethod
    def create_event_row(start_parsed, end_parsed, duration_event, week_nr, week_dur, event_summary):
        duration_str = f"{duration_event.total_seconds() // 3600:02.0f}:{(duration_event.total_seconds() // 60) % 60:02.0f}"
        week_dur_str = f"{week_dur.total_seconds() // 3600:02.0f}:{(week_dur.total_seconds() // 60) % 60:02.0f}"
        return pd.DataFrame({
            'Date': [start_parsed.strftime("%d-%m-%Y")],
            'Day': [start_parsed.strftime("%d")],
            'Start_time': [start_parsed.strftime("%H:%M")],
            'End_time': [end_parsed.strftime("%H:%M")],
            'Duration': [duration_event],  # Store as timedelta object
            'Week_nr': [week_nr],
            'Week_duration': [week_dur_str],
            'Description': [event_summary]
        })

    def generate_timesheet(self, start_date, end_date, week_totals=False, output_format: OutputFormat = OutputFormat.TABLE, selected_clients=None):
        self.output_format = output_format
        
        # Set logging level based on output format
        if output_format in [OutputFormat.CSV, OutputFormat.TOTAL]:
            logging.getLogger().setLevel(logging.WARNING)
        else:  # TABLE format
            logging.getLogger().setLevel(logging.INFO)
        
        events = self.get_gcal_events(start_date, end_date)
        if not events:
            logging.info("No events found between %s and %s", start_date, end_date)
            return []

        client_list_to_process = selected_clients if selected_clients else self.client_list
        time_sheets = self.process_events(events, client_list_to_process)

        for sheet in time_sheets:
            self.add_totals_to_sheet(sheet, week_totals)
            self.print_sheet_summary(sheet, output_format)

        return time_sheets

    def add_totals_to_sheet(self, sheet, week_totals):
        sheet.time_sheet_df.insert(loc=6, column="Week_total", value="")
        sheet.time_sheet_df.insert(loc=1, column="Day_total", value="")

        # Duration is already a timedelta, no need to convert

        if week_totals:
            week_tot_grp = sheet.time_sheet_df.groupby(["Week_nr"])["Week_duration"].last()
            for index, row in sheet.time_sheet_df.iterrows():
                if row["Week_nr"] in week_tot_grp.index:
                    sheet.time_sheet_df.at[index, "Week_total"] = week_tot_grp[row["Week_nr"]]

        day_tot_grp = sheet.time_sheet_df.groupby(["Week_nr", "Day"])["Duration"].sum()
        sheet.time_sheet_df["Day_total"] = sheet.time_sheet_df.groupby(["Week_nr", "Day"])["Duration"].transform('sum')
        sheet.time_sheet_df["Day_total"] = sheet.time_sheet_df["Day_total"].apply(
            lambda x: f"{x.total_seconds() // 3600:02.0f}:{(x.total_seconds() // 60) % 60:02.0f}")

        # Convert Duration to string format for display
        sheet.time_sheet_df["Duration"] = sheet.time_sheet_df["Duration"].apply(
            lambda x: f"{x.total_seconds() // 3600:02.0f}:{(x.total_seconds() // 60) % 60:02.0f}")

    def print_sheet_summary(self, sheet, output_format: OutputFormat):
        if output_format == OutputFormat.TOTAL:
            total_hours = sheet.total_duration.total_seconds() / 3600
            print(f"{total_hours:.2f}")
        elif output_format == OutputFormat.CSV:
            print(sheet.time_sheet_df.to_csv(index=False))
        else:  # TABLE
            print(f"\nTime sheet for client: {sheet.client_name}")
            total_hours, remainder = divmod(sheet.total_duration.total_seconds(), 3600)
            total_minutes = remainder // 60
            print(f"Total duration for client was: {total_hours:.0f} hours and {total_minutes:.0f} minutes.")
            print(tabulate(sheet.time_sheet_df, headers=sheet.time_sheet_df.columns, tablefmt="presto"))


def main():
    parser = argparse.ArgumentParser(description="Generate time sheets from Google Calendar events.")
    parser.add_argument("-s", "--start", help="Start date (format: DD/MM/YYYY)", type=str)
    parser.add_argument("-e", "--end", help="End date (format: DD/MM/YYYY)", type=str)
    parser.add_argument("-l", "--last", action="store_true", help="Generate time sheet for last month")
    parser.add_argument("-t", "--this", action="store_true", help="Generate time sheet for this month")
    parser.add_argument("-w", "--weektotals", action="store_true", help="Include week totals in the report")
    parser.add_argument("-lc", "--list-calendars", action="store_true", help="List available calendars")
    parser.add_argument("-f", "--format", 
                       type=OutputFormat, 
                       choices=list(OutputFormat), 
                       default=OutputFormat.TABLE,
                       help="Output format (table, csv, or total)")
    args = parser.parse_args()

    generator = TimesheetGenerator()

    if args.list_calendars:
        generator.list_calendars()
        return
    if args.start and args.end:
        start_date = parse(args.start, dayfirst=True)
        end_date = parse(args.end, dayfirst=True).replace(hour=23, minute=59, second=59)
    elif args.this:
        start_date = generator.first_day_this_month
        end_date = generator.last_day_this_month
    elif args.last:
        start_date = generator.first_day_last_month
        end_date = generator.last_day_last_month
    else:
        start_date = generator.first_day_last_month
        end_date = generator.last_day_last_month

    generator.generate_timesheet(start_date, end_date, args.weektotals, args.format)


if __name__ == '__main__':
    main()