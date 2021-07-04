# TimeSheeterPython
Generates a timesheet based on Google Calendar events that contain "@clientname" in their description.

# About:
This script downloads Google calendar events, finds all events with a "@clientname" tag and generates a 
time sheet table with: 
`| Day of month | Start_time | End_time | Duration | Week_nr | Week_duration | Description`

The output can be a ascii table or a pdf based on a .docx template (work in progress).
An example of the output: 

![Example output](./../images/example_output1.png)

# Usage: 

`
python TimeSheeter.py -h
usage: TimeSheeter.py [-h] [-s START] [-e END] [-l] [-t] [-w] [-r]

optional arguments:
  -h, --help            show this help message and exit
  -s START, --start START
                        Start date. Script will assume time of day 0:00:00 unless otherwise specified.
  -e END, --end END     End date. Script will assume time of day 23:59:59 unless otherwise specified.
  -l, --last            Generate time sheet for all of last month.
  -t, --this            Generate time sheet for all of this month.
  -w, --weektotals      Add week numbers and totals per week to report.
  -r, --report          Generate a .pdf report based on the .docx template file in this directory.
`

#####TODO:
* better auth flow without requiring dev console stuff
* general code cleanup
* add option to output straight to HTML, Word or Excel (with markup?)

