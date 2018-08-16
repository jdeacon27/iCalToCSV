# Process an iCal file, extracting fields, beginning with start date and description.
#
# Originally this was intended for converting a Google calendar export
# into CSV for import into a "dear diary" spreadsheet. It was extended to process
# more of the available fields.
#
# For some reason, gcal adds newlines at the start of descriptions.
# This strips them out. If ever anyone else uses this, perhaps that should become
# optional.

# Change log:
# 07-Aug-18 Add processing of start dates with times. These have a different
#			format to "All day" events. The only thing that's done with them
#			is to emit a warning; otherwise the date portion is processed as
#			normal.
# 08-Aug-18	Add acceptance of bundled field arguments, e.g. -sed.
# 09-Aug-18	Move printing of date and description into the processing of END:VEVENT
# 09-Aug-18	Begin processing of field selection switches
# 09-Aug-18	Change to using a hash of references to call the column output functions
# 10-Aug-18	Populate function despatch array
# 10-Aug-18	Unfortunately, it appears that whoever it was was wrong when they said
#			that Perl hashes preserve declaration order. I'm seeing an apparently random order.
#			So I worked out how to install CPAN modules and tried Tie::IxHash again.
# 11-Aug-18	Add routines for extracting and emitting other fields
# 11-Aug-18 More entries have turned up where Google adds multiple \n at the start of
#			the descrition. I've modified the description's RE.
# 11-Aug-18	Change processDate() to return its results.
# 12-Aug-18	Add rest of seemingly intersting fields. It's possible that people
#			will find other iCal fields that are of interest.
# 15-Aug-18	Change from warn function for debug to conditional STDLOGing
# 16-Aug-18	Correct usage to match github README.MD
# 16-Aug-18 Change a complicated 'if' into an 'unless'; and correct processing of --help option

use strict;
use Data::Dumper;
use Getopt::Long;
use Tie::IxHash;

my $false = 0; my $true = 1;
my $debug = $false;
my $usageString = "\nUsage: iCalToCSV.pl [--help] [-ysteudla] --start=STARTDATE --end=ENDDATE <INFILE >OUTFILE\n".
			"The STARTDATE and ENDDATE parameters should be given in reverse date style: yyyymmdd e.g.:\n".
			"--start=20171225\n".
			"and where y,s,e, etc. are switches for inclusion of the various iCalendar fields:\n".
			"y = Summary (Event name)\ns = Start Date\ne = End Date\nt = Start Time\nu = EndTime\nd = Description\n".
			"l = Location\na = Status\n";
my $helpString = "Converts iCal to CSV (comma separated values), for subsequent import into something like Excel.\n\n".
			"It doesn't check the calendar (which would not be difficult to add) and will export ".
			"all events within the date range.\n\n".
			"The date range is the only way to select the entries that are processed and exported, and those arguments are mandatory.".
			"Within the events selcted by the date range parameter, you can choose the fields to ".
			"be exported using the -ysteud etc. arguments. If no -ysteud etc. arguments are supplied, the ".
			"default is to export start date and description (-sd).\n\n".
			"This was originally written to process diary entries exported from Google Calendar; ".
			"diary entries which were \"All Day\" events. iCalendar has two DATE styles, one of ".
			"which is untimed. That's what Google Calendar uses and this is detected. ".
			"If you elect to print the time of an All Day event, you'll get 0:0:0. Google, however, ".
			"also sets an End Date on All Day events. End dates on All Day events are ignored; you'll ".
			"get blanks in the output.";

my $processingEvent = $false;
my $processingDescription = $false;
my $foundStartDate = $false;
my $foundUntimedStartDate = $false;
my $dateInPeriod = $false;
my $summary = "";
my $startDate = "";
my $startTime = "";
my $endDate = "";
my $endTime = "";
my $description = "";
my $location = "";
my $status = "";
my %fields;
my $fieldsAsTiedHash = tie(%fields, 'Tie::IxHash');	# A (wrapper?) around the Perl hash to preserve order
my $fieldSwitches = "";
my $numberOfFieldsWanted = 0;
my $summaryWanted = $false;
my $startDateWanted = $false;
my $endDateWanted = $false;
my $descriptionWanted = $false;
my $startTimeWanted = $false;
my $endTimeWanted = $false;
my $locationWanted = $false;
my $statusWanted = $false;
my @functionRefs;
my $startOfPeriod = 0;
my $endOfPeriod = 0;

# Subroutines
sub processDate {
	my ($dayNumber, $monthNumber, $yearNumber, $wholeDateNumber, $hoursNumber, $minutesNumber, $secondsNumber) = @_;
	my $date; my $time;
	if ( $wholeDateNumber >= $startOfPeriod && $1 <= $endOfPeriod ) {
		$dateInPeriod = $true;
	}
	$date = "$dayNumber\/$monthNumber\/$yearNumber" if $dateInPeriod;
	$time = "$hoursNumber:$minutesNumber:$secondsNumber" if $dateInPeriod;
	return($date, $time);
}

sub parseArguments {
	Getopt::Long::Configure ("bundling");
	GetOptions (\%fields, 'start=i', 'end=i', 'y', 's', 'e', 't', 'u', 'd', 'l', 'a', 'help');
	unless ( ((exists $fields{start}) && (exists $fields{end})) ||
				exists $fields{help} ) {
		die $usageString;
	}

	$startOfPeriod = $fields{start};
	$endOfPeriod = $fields{end};

	# Create one string with the single -sed etc. arguments concatenated and, hopefully, in order
	foreach ( keys %fields ) {
if ( $debug ) { print STDERR "$_ = $fields{$_}\n"; }
		$fieldSwitches .= $_ if length == 1; 	# length defaults to taking $_ as its argument
	}
	$numberOfFieldsWanted = length($fieldSwitches);
if ( $debug ) { print STDERR "\$fieldSwitches = $fieldSwitches"; }
	# TODO Will the following become redundant when we process the loop based on the array of references built up from the field switches?
	$summaryWanted = $true if exists $fields{y};
	$startDateWanted = $true if exists $fields{s};
	$endDateWanted = $true if exists $fields{e};
	$startTimeWanted = $true if exists $fields{t};
	$endTimeWanted = $true if exists $fields{u};
	$descriptionWanted = $true if exists $fields{d};
	$locationWanted = $true if exists $fields{l};
	$statusWanted = $true if exists $fields{a};
	# Add more if exists for more switches when added

	# Set defaults for absent field switches
	if ( $numberOfFieldsWanted == 0 ) {	# field switches were not supplied so default to date and description
		$startDateWanted = $true;
		$descriptionWanted = $true;
		$fieldSwitches = "sd";
		$numberOfFieldsWanted = 2;
	return;
	}
}

sub orderColumns {
	# A bit inelegant. Prints out the column headings; AND populates the function references array
	my $field;
	for ($field = 0; $field < $numberOfFieldsWanted; $field++) {
		if ( substr($fieldSwitches, $field, 1) eq 'y' ) {
			print "Summary";
			$functionRefs[$field] = \&emitSummary;
		}
		if ( substr($fieldSwitches, $field, 1) eq 's' ) {
			if ( $endDateWanted ) {
				print "Start Date";
			} else {
				print "Date";
			}
			$functionRefs[$field] = \&emitDate;
		}
		if ( substr($fieldSwitches, $field, 1) eq 'e' ) {
			print "End Date";
			$functionRefs[$field] = \&emitEndDate;
		}
		if ( substr($fieldSwitches, $field, 1) eq 't' ) {
			print "Start Time";
			$functionRefs[$field] = \&emitStartTime;
		}
		if ( substr($fieldSwitches, $field, 1) eq 'u' ) {
			print "End Time";
			$functionRefs[$field] = \&emitEndTime;
		}
		if ( substr($fieldSwitches, $field, 1) eq 'd' ) {
			print "Description";
			$functionRefs[$field] = \&emitDescription;
		}
		if (substr($fieldSwitches, $field, 1) eq 'l' ) {
			print "Location";
			$functionRefs[$field] = \&emitLocation;
		}
		if (substr($fieldSwitches, $field, 1) eq 'a' ) {
			print "Status";
			$functionRefs[$field] = \&emitStatus;
		}
		if ( ($field < $numberOfFieldsWanted - 1) ) {		# Because CSV is comma SEPARATED not TERMINATED
			print ",";
		}
	}
	print "\n";
if ( $debug ) { print STDERR "\$#functionRefs = @{[$#functionRefs + 1]}\n"; }	# That's a dense trick to perform arithmetic in an interpolated string. See https://stackoverflow.com/questions/3939919/can-perl-string-interpolation-perform-any-expression-evaluation
	return;
}

sub emitSummary {
	print "\"";
	print "$summary";
	print "\"";
	return;
}

sub emitDate {
	print "$startDate";
	return;
}

sub emitEndDate {
	print "$endDate";
	return;
}

sub emitStartTime {
	print "$startTime";
	
	return;
}

sub emitEndTime {
	print "$endTime";
	return;
}

sub emitDescription {
	print "\"";
	print $description;
	print "\"";
	return;
}

sub emitLocation {
	print "$location";
	return;
}

sub emitStatus {
	print "$status";
	return;
}

sub callFieldEmitters {
	# Use the array of function reference, built from the -setudla switches, to
	# call the emit functions in supplied order of switches
	my $index;
	if ( $dateInPeriod ) {
		for ($index = 0; $index < $numberOfFieldsWanted; $index++) {
if ( $debug ) { print STDERR "\$index = $index\n"; }
			$functionRefs[$index]->();
			print "," unless ( $numberOfFieldsWanted == 1 || $index == $numberOfFieldsWanted-1 );
		}
		print "\n";
	}
	$processingEvent = $false;
	$processingDescription = $false;
	$foundStartDate = $false;
	$foundUntimedStartDate = $false;
	$dateInPeriod = $false;
	return;
}
# End subroutines


# Main
parseArguments();
if ( exists $fields{help} ) {
	warn "$usageString\n";
	warn "$helpString\n";
	die;
}
orderColumns();
# Set up the hash of funtion references
# $functionRefs[0] = \&emitDate;
# $functionRefs[1] = \&emitDescription;

if ( $debug ) { print STDERR "\$fieldSwitches = $fieldSwitches"; }
if ( $debug ) { print STDERR "\$numberOfFieldsWanted = $numberOfFieldsWanted"; }

# Loop
while( <> ) {
	if ( /BEGIN:VEVENT/ ) {
		warn " WARNING: VEVENT BEGIN encountered while already processing VEVENT" if $processingEvent;
		$processingEvent = $true;
	}
	if ( /END:VEVENT/ ) {
		warn " WARNING: VEVENT END encountered when not processing VEVENT" if !$processingEvent;
		callFieldEmitters();
	}
	if ( $processingDescription ) {
		if ( /^ (.+)/ ) {
			$description .= $1;
		} else {
			$processingDescription = $false;
			$description =~ s/\\n/\n/g;	# swap the two-character pair '\' and 'n' for a newline
			$description =~ s/\\,/,/g;	# comma doesn't need escapting in CSV
			$description =~ s/\\;/;/g;	# semi-colon doesn't need escapting in CSV
			$description =~ s/\"/\"\"/g;	# double quote marks do need escaping however
		}
	}

	# The way Google calendar exports "all day" events is to use two DT*VALUE=DATE entries
	# with the END set to the next day, so we have to assume and note an "all day" and
	# thereupon ignore the end date, with $foundUntimedStartDate
	
	if ( /DTSTART:((\d{4})(\d{2})(\d{2}))T(\d{2})(\d{2})(\d{2}).*/ ) {
		warn " WARNING: Entry with a time. Forget to make an entry All Day? $4\/$3\/$2\n";
		# warn " WARNING: Found DATE more than once" if $foundStartDate;
		$foundStartDate = $true;
		($startDate, $startTime) = processDate($4, $3, $2, $1, $5, $6, $7);
	}

	if ( /DTSTART;VALUE=DATE:((\d{4})(\d{2})(\d{2}))/ ) {	# NB nested backreferences
		# warn " WARNING: Found DATE more than once" if $foundStartDate;
		$foundStartDate = $true;
		$foundUntimedStartDate = $true;
		($startDate, $startTime) = processDate($4, $3, $2, $1, 00, 00, 00);
	}

	if ( /DTEND:((\d{4})(\d{2})(\d{2}))T(\d{2})(\d{2})(\d{2}).*/ ) {
		($endDate, $endTime) = processDate($4, $3, $2, $1, $5, $6, $7);
	}

	if ( /DTEND;VALUE=DATE:((\d{4})(\d{2})(\d{2}))/ ) {	# NB nested backreferences
		($endDate, $endTime) = processDate($4, $3, $2, $1, 00, 00, 00) unless $foundUntimedStartDate;
	}

	if ( /DESCRIPTION:(\\n)*(.*)/ ) {
		if ( $dateInPeriod ) {
			# This is a bit lonely now that more fields are processed. See Notes
			warn " WARNING: Found DESCRIPTION without having found DATE" if !$foundStartDate;
			$processingDescription = $true;
			if ( $2 eq '' ) {
				$description = "EMPTY DESCRIPTION";
			} else {
				$description = $2;
			}
		}
	}
	if ( /SUMMARY:(.*)/ ) {
		if ( $dateInPeriod ) {
			if ( $1 eq '' ) {
				$summary = "EMPTY SUMMARY";
			} else {
				$summary = $1;
			}
			# This is a duplicate of (multi-line) description massaging. Consider using a subroutine
			$summary =~ s/\\n/\n/g;		# swap the two-character pair '\' and 'n' for a newline
			$summary =~ s/\\,/,/g;		# comma doesn't need escapting in CSV
			$summary =~ s/\\;/;/g;		# semi-colon doesn't need escapting in CSV
			$summary =~ s/\"/\"\"/g;	# double quote marks do need escaping however
		}
	}
	if ( /LOCATION:(.*)/ ) {
		if ( $dateInPeriod ) {
			if ( $1 eq '' ) {
				$location = "EMPTY LOCATION";
			} else {
				$location = $1;
			}
		}
	}
	if ( /STATUS:(.*)/ ) {
		if ( $dateInPeriod ) {
			if ( $1 eq '' ) {
				$status = "EMPTY STATUS";
			} else {
				$status = $1;
			}
		}
	}
}

# Notes
# I could construct a string with the argument letters in order of declaration,
# if Perl hashes do indeed preserve the order of declaration.
# Then I could use the characters of that string, in order, to despatch to the
# various column printing routines, probably using function references.
#
# Currently assumes that the SUMMARY (what you or I would call the Name of the entry)
# takes a single line in the iCal file. If it can be multiple lines (indeed any other
# fields that turn out to be multi-line) will need handling in the same way as
# description is handled.
#
# Currently assumes that only SUMMARY and DESCRIPTION might have embedded commas.
# If other fiels could, then they need quoting in the CSV as well.
# Well, presumably LOCATION could at least.
# Could consider quoting every field, but that would require escaping real quotes.
#
# There is a check for an entry with a description but no date. This was when we
# were purely processing diary-style entries. Now we're processing more fields,
# this needs thinking about a bit more.
#
# Also, as I was only processing all-day events, I didn't concern myself with
# time formats. Google export uses one of three possibly iCalendar formats: the
# UTC format. Everything is Zulu. That's the T999999Z kind of entry.
#
# Could add a switch --checkForEmpty and only put in all those "EMPTY DESCRIPTION"s
# etc. if the switch is present.
