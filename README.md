# TimeSheeterPython
Generates a timesheet based on Google Calendar events that contain "@tagname" in their description

# About:
# This script downloads Google calendar events, finds all events with a "@something" tag (the tag is defined in the python code)
# and generates a time sheet table with:
# total duration, day, start and end times, event duration and description.
# Currently the table is tab seperated so it can be pasted into some spreadsheet software. 

#####TODO:
# better auth flow without requiring dev console stuff
# break up into nice functions. Load calendar events into Python object with proper types and string equivalents?
# accept arguments when running python file e.g. "Python TimeSheeter.py -getLastMonth -clientTag -clientName"
# add option to output straight to HTML, Word or Excel (with markup?)
#
