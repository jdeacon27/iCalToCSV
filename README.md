# ical2csv
Convert iCalender (files) to CSV (files) - Perl script

There are other utilities for converting iCal to CSV, but none of them did what I want or I couldn't work out how to use them. This one at least has a help option: perl iCalToCSV --help

As far as I'm aware, none of the others are written in Perl; which might be very sensible on their part.

Extract from the help:
Converts iCal to CSV (comma separated values), for subsequent import into something like Excel.

It doesn't check the calendar (which would not be difficult to add) and will export all events within the date range. A date range is the only way to select the entries that are processed and exported, and the date range arguments are mandatory.

Within the events selcted by the date range parameter, you can choose the fields to be exported using the -ysteud etc. arguments. If no -ysteud etc. arguments are supplied, the default is to export start date and description (-sd).

This was originally written to process diary entries exported from Google Calendar; diary entries which were "All Day" events. iCalendar has two DATE styles, one of which is untimed. That's what Google Calendar uses for All Day events and this is detected. If you elect to include the time of an All Day event, you'll get 0:0:0. Google, however, also sets an End Date on All Day events. End dates on All Day events are ignored; if you elect to include the end of an All Day event, you'll get blanks in the output.

Usage:
iCalToCSV.pl [--help] [-ysteudla] --start=STARTDATE --end=ENDDATE \<INFILE \>OUTFILE

The STARTDATE and ENDDATE parameters should be given in reverse date style: yyyymmdd, e.g.:

	--start=20171225

and where y,s,e,t, etc. are switches for inclusion of the various iCalendar fields:

	y = Summary (Event name)
	s = Start Date
	e = End Date
	t = Start Time
	u = EndTime
	d = Description
	l = Location
	a = Status
