#!/usr/bin/perl
# Perl - v: 5.16.3
#------------------------------------------------------------------------------#
# Tool name   : RDP-Parser
# Website     : http://le-tools.com/
# SourceForge	: https://sourceforge.net/p/rdp-parser
# GitHub		  : https://github.com/arioux/rdp-parser
# Description : RDP-Parser extracts RDP activities from Microsoft Windows Event Logs
# Creation    : 2018-08-09
# Modified    : 2018-09-19
my $VERSION   = "1.0";
# Author      : Alain Rioux (admin@le-tools.com)
#
# Copyright (C) 2018  Alain Rioux (le-tools.com)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#------------------------------------------------------------------------------#

#------------------------------------------------------------------------------#
# Modules
#------------------------------------------------------------------------------#
use strict;
use warnings;
use Cwd;
use Filter::Arguments;
use File::Copy;
use Time::Local 'timelocal';
use Parse::EventLog;
use Win32::EventLog;
use Win32::RunAsAdmin;
use Excel::Writer::XLSX;

#------------------------------------------------------------------------------#
# Global variables
#------------------------------------------------------------------------------#
my $PROG = $0;
my $PROGDIR = getdcwd;                                                         # Program path
#while (chop($PROGDIR) ne "\\") { }                                            # Dir only
my $LOG_PATH	= Argument( alias => 'p', default => 'current' );								 # Options: Full path or current
my $TYPE			= Argument( alias => 't', default => '1' );											 # 1: minimal (default, event with IP addresses only, less event details)
																																							 # 2: minimal with IP addresses
																																							 # 3: normal (event with public IP addresses only)
																																							 # 4: normal with all IP addresses
																																							 # 5: full (all RDP events)
my $DATE_START		= Argument( alias => 's' );																	 # [format: yyyy-mm-dd] default is none
my $DATE_END			= Argument( alias => 'e' );																	 # [format: yyyy-mm-dd] default is none
my $REPORT_FORMAT	= Argument( alias => 'r', default => '1');									 # Options are: 1: xlsx, 2: text, 3: html
my $DATA_STR			= Argument( alias => 'l' );																	 # Data strings on a single line
my $OPEN_REPORT		= Argument( alias => 'o');																	 # Open report at the end
my $BACKUP				= Argument( alias => 'b');																	 # Backup all Event logs from live system
my $HELP					= Argument( alias => 'h');																	 # Print Help
my $AS_ADMIN			= Argument( alias => 'asadmin');														 # Restarted with admin rights

#------------------------------------------------------------------------------#
# Starting program
#------------------------------------------------------------------------------#

if ($HELP) {
	my	$menu	 = "\nRDP-Parser $VERSION\n";
			$menu .= "***********************************************************************\n";
			$menu .= "Copyright (C) 2018 Alain Rioux (le-tools.com). All rights reserved.\n\n";
			$menu .= "This program is free software: you can redistribute it and/or modify\n";
			$menu .= "it under the terms of the GNU General Public License as published by\n";
			$menu .= "the Free Software Foundation, either version 3 of the License, or\n";
			$menu .= "(at your option) any later version.\n\n";
			$menu .= "This program is distributed in the hope that it will be useful,\n";
			$menu .= "but WITHOUT ANY WARRANTY; without even the implied warranty of\n";
			$menu .= "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the\n";
			$menu .= "GNU General Public License for more details.\n\n";
			$menu .= "You should have received a copy of the GNU General Public License\n";
			$menu .= "along with this program.  If not, see <http://www.gnu.org/licenses/>.\n";
			$menu .= "***********************************************************************\n";
			$menu .= "usage: RDP-Parser [options]\n";
			$menu .= "Options and arguments:\n";
			$menu .= "--p\t: Path: default is current or ".'C:\Windows\System32\winevt\Logs'."\n";
			$menu .= "--t\t: Type: 1: minimal (default, event with IP addresses only, less event details)\n";
			$menu .= "\t  2: minimal with all IP addresses\n";
			$menu .= "\t  3: normal (event with IP addresses only)\n";
			$menu .= "\t  4: normal with all IP addresses\n";
			$menu .= "\t  5: full (all RDP and login events)\n";
			$menu .= "--s\t: Date start: [format: yyyy-mm-dd]\n";
			$menu .= "--e\t: Date end: [format: yyyy-mm-dd]\n";
			$menu .= "--r\t: Report format: 1: xlsx (default), 2: text, 3: html\n";
			$menu .= "--l\t: Data strings on a single line\n";
			$menu .= "--o\t: Open report at the end\n";
			$menu .= "--b\t: Copy all Event logs from live system.\n";
			$menu .= "--h\t: Print this help message and exit\n";
			$menu .= "Examples:\n";
			$menu .= "- RDP-Parser (without any argument): Print *minimal* type for Event Logs in current\n";
			$menu .= "  dir or system\n";
			$menu .= "- RDP-Parser --t 2 --s 2018-01-01 --e 2019-01-01: Print *normal* type for Event\n";
			$menu .= "  Logs in current dir or system, all events in 2018\n";
			$menu .= "***********************************************************************\n\n";
	print $menu;
	exit(0);
} elsif ($BACKUP) {
	# Request admin rights to copy Event Logs in the current directory
	if (not Win32::RunAsAdmin::check) {
		if ($AS_ADMIN) {
			print "This function requires admin rights.\n";
			exit(0);
		}
		print "Admin rights required to copy live Event Logs.\n";
		my $params;
		foreach (@ARGV) { $params .= $_ . ' '; }
		$params .= '--asadmin';
		print "Program will restart in a new console.\n";
		Win32::RunAsAdmin::run($PROG, $params, $PROGDIR);
		exit(0);
	} else {
		# copy all Event Logs files in the current directory
		my $systemELDir = "$ENV{'WINDIR'}\\Sysnative\\winevt\\Logs";
		if (opendir(DIR,"$systemELDir\\")) {
			my $logDir = "$PROGDIR\\Logs";
			mkdir($logDir) if !-d $logDir;
			my $i = 0;
			print "Event Logs will be copied from live system to current directory...\n";
			while (my $file = readdir(DIR)) {
				if ($file =~ /\.evtx$/) {
					print "Copying $file... ";
					my $filePath = "$systemELDir\\$file";
					if (copy($filePath, "$logDir\\$file")) {
						print "Copied.\n";
						$i++;
					} else { print "Error.\n"; }
				}
			}
			close(DIR);
			print "\n$i Event logs have been copied. Program exits.\n";
		} else { print "Cannot open Event logs directory.\n"; }
		sleep(10);
		exit(0);
	}
} else {
	print "\nRDP-Parser $VERSION\n";
	print "***********************************************************************\n";
	foreach (@ARGV) { print "Warning: $_ is not a valid option. Do you mean -$_?" if (/^\-[^-]/); }
	my ($elSecFile, $elTSLocalFile, $elTSRemoteFile, $elSysFile);
	# Verify path: specified path, current if there are evtx or C:\Windows\System32\winevt\Logs if access allowed
	if ($LOG_PATH eq 'current') {
		$LOG_PATH = $PROGDIR;
		($elSecFile, $elTSLocalFile, $elTSRemoteFile, $elSysFile) = &listEventLogs($LOG_PATH);
		if (!$elSecFile and !$elTSLocalFile and !$elTSRemoteFile and !$elSysFile) {
			my $answer;
			if (!$AS_ADMIN) {
				print "No log found in current directory and no path specified. Do you want to\n" .
							"parse Event Logs from current system (admin rights required)?\n" .
							"[Type (y)es to continue]: ";
				$answer = <STDIN>;
				chomp($answer);
			} else { $answer = 'y'; } # Tool has been restarted with admin rights
			if ($answer =~ /ye?s?/) {
				# Request admin rights to copy Event Logs in the current directory
				if (not Win32::RunAsAdmin::check) {
					print "Admin rights required to copy live Event Logs.\n";
					my $params;
					foreach (@ARGV) { $params .= $_ . ' '; }
					$params .= '--asadmin';
					print "Program will restart in a new console.\n";
					Win32::RunAsAdmin::run($PROG, $params, $LOG_PATH);
					exit(0);
				} else {
					# copy Event Logs in the current directory
					my $systemELDir = "$ENV{'WINDIR'}\\Sysnative\\winevt\\Logs";
					# Copy each Event Logs file as it is not possile to open Event Logs on live system (require admin rights)
					print "Event Logs must be copied from live system to current directory...\n";
					if (-d $systemELDir and
							(-e "$systemELDir\\Security.evtx" and copy("$systemELDir\\Security.evtx", "$LOG_PATH\\Security.evtx")) and
							(-e "$systemELDir\\Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx" and
							 copy("$systemELDir\\Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx",
										"$LOG_PATH\\Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx")) and
							(-e "$systemELDir\\Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx" and
							 copy("$systemELDir\\Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx",
										"$LOG_PATH\\Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx")) and
							(-e "$systemELDir\\System.evtx" and copy("$systemELDir\\System.evtx", "$LOG_PATH\\System.evtx"))) {
						print "Event Logs have been copied to the current directory.\n";
					} else { print "Error: Cannot gather Event Logs from system.\n"; }
				}
			} else { exit(0); }
		}
	}
	# List Event logs in directory
	if (-d $LOG_PATH) {
		if (!$elSecFile and !$elTSLocalFile and !$elTSRemoteFile and !$elSysFile) {
			($elSecFile, $elTSLocalFile, $elTSRemoteFile, $elSysFile) = &listEventLogs($LOG_PATH);
			if (!$elSecFile and !$elTSLocalFile and !$elTSRemoteFile and !$elSysFile) {
				print "No Event logs in the given path or access to path is disallowed (use --h)\n";
				exit(0);
			}
		}
	} else {
		print "Not a valid path or access to path is disallowed (use --h)\n";
		exit(0);
	}
	# Verify type: minimal, normal or all
	if ($TYPE !~ /^[1-5]$/) {
		print "Valid [type] are 1: minimal (default), 2: normal and 3: all (use --h)\n";
		exit(0);
	}
	# Verify given dates
	if ($DATE_START) {
		if ($DATE_START =~ /^\d{4}\-\d{2}\-\d{2}$/) {
			my ($y, $m, $d) = split(/\-/, $DATE_START);
			$m--;
			$DATE_START = timelocal(0,0,0,$d,$m,$y); # Store in Unixtime format
		} else {
			print "Invalid starting date (use --h)\n";
			exit(0);
		}
	}
	if ($DATE_END) {
		if ($DATE_END =~ /^\d{4}\-\d{2}\-\d{2}$/) {
			my ($y, $m, $d) = split(/\-/, $DATE_END);
			$m--;
			$DATE_END = timelocal(0,0,0,$d,$m,$y); # Store in Unixtime format
			$DATE_END += 86400; # Add one day, because last day is included
		} else {
			print "Invalid ending date (use --h)\n";
			exit(0);
		}
	}
	# Verify report
	if ($REPORT_FORMAT !~ /^[123]$/) {
		print "Valid [report] are 1: xlsx (default), 2: text and 3: html (use --h)\n";
		exit(0);
	}
	# Logs:
	#		$elSecFile (Security.evtx)
	# 	$elTSLocalFile (Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx)
	# 	$elTSRemoteFile (Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx)
	# 	$elSysFile (System.evtx)
	my %Events;
	# Events:
	#		RecordNumber
	#		TimeGenerated
	#		Timewritten
	#		Computer
	#		Source
	#		Category
	#		EventID
	#		EventType
	#		Details
	
	# Parse each Event log file
	my $nbrActivities = 0;
	if ($elSecFile and ($elSecFile =~ /security.evt$/ or $elSecFile =~ /SecEvent.Evt$/)) { # Old format
		$nbrActivities += &parseOldEventLog($elSecFile, \%Events);
	} else { # Evtx format
		$nbrActivities += &parseEventLog(1, $elSecFile			, \%Events)	if $elSecFile;
		$nbrActivities += &parseEventLog(2, $elTSLocalFile	, \%Events)	if $elTSLocalFile;
		$nbrActivities += &parseEventLog(3, $elTSRemoteFile	, \%Events)	if $elTSRemoteFile;
		$nbrActivities += &parseEventLog(4, $elSysFile			, \%Events)	if $elSysFile;
	}	
	# Produce report
	if ($nbrActivities) {
		my $dateStr	= &date(time);
		$dateStr		=~ s/:/-/g;
		$dateStr		=~ s/ /_/g;
		my $report	= 'report_' . $dateStr;
		if 		($REPORT_FORMAT == 1)	{ $report .= '.xlsx'; }
		elsif ($REPORT_FORMAT == 2)	{ $report .= '.txt';	}
		else 												{ $report .= '.html'; }
		print "$nbrActivities activities have been found.\nCreating $report...\n";
		if 		($REPORT_FORMAT == 1)	{ &reportXLSX($report, \%Events); }
		elsif ($REPORT_FORMAT == 2)	{ &reportTXT( $report, \%Events); }
		else 												{ &reportHTML($report, \%Events); }
		if ($OPEN_REPORT and $report and -e $report) {
			print "Opening $report...\n";
			system("cmd /c start $report") ;
		} else { print "Report has been created.\n"; }		
	} else { sleep(10); print "No RDP activity found.\n"; sleep(10); }
	exit(0);
}

#--------------------------#
sub listEventLogs
#--------------------------#
{
	my $path = shift;
	my ($elSecFile, $elTSLocalFile, $elTSRemoteFile, $elSysFile) = undef;
	if (opendir(DIR,"$path\\")) {
		print "Listing Event Logs in $path...\n";
		while (my $file = readdir(DIR)) {
			my $filePath = "$path\\$file";
			if			($file eq 'Security.evtx') {
				$elSecFile = $filePath;
				print "Security.evtx found.\n";
			} elsif ($file eq 'Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx') {
				$elTSLocalFile = $filePath;
				print "Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx found.\n";
			} elsif ($file eq 'Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx') {
				$elTSRemoteFile = $filePath;
				print "Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational found.\n";
			} elsif ($file eq 'System.evtx') {
				$elSysFile = $filePath;
				print "System.evtx found.\n";
			} elsif ($file =~ /security.evt$/) {
				$elSecFile = $filePath;
				print "security.evt found.\n";
			} elsif ($file =~ /SecEvent.Evt$/) {
				$elSecFile = $filePath;
				print "SecEvent.Evt found.\n";
			}
		}
		closedir(DIR);
		print "\n";
	}
	return($elSecFile, $elTSLocalFile, $elTSRemoteFile, $elSysFile);
	
}   #--- End listEventLogs

#--------------------------#
sub parseEventLog
#--------------------------#
{
	# Local variables
	my ($elog, $file, $refEvents) = @_;
	# Elog: 1 = Security.evtx,
	#				2 = Microsoft-Windows-TerminalServices-LocalSessionManager%4Operational.evtx,
	#				3 = Microsoft-Windows-TerminalServices-RemoteConnectionManager%4Operational.evtx,
	#				4 = System.evtx
	my %eventIDs;
	$eventIDs{21} 	= 'Remote Desktop Services: Session logon succeeded';
	$eventIDs{22} 	= 'Remote Desktop Services: Shell start notification received';
	$eventIDs{23} 	= 'Remote Desktop Services: Session logoff succeeded';
	$eventIDs{24} 	= 'Remote Desktop Services: Session has been disconnected';
	$eventIDs{25} 	= 'Remote Desktop Services: Session reconnection succeeded';
	$eventIDs{39} 	= 'Session <X> has been disconnected by session <Y>';
	$eventIDs{40} 	= 'Session <X> has been disconnected, reason code <Z>';
	$eventIDs{56} 	= 'The Terminal Server security layer detected an error in the protocol stream and has disconnected the client.';
	$eventIDs{261}	= 'Listener RDP-Tcp received a connection';
	$eventIDs{1149} = 'User authentication succeeded';
	$eventIDs{4624} = 'An account was successfully logged on';
	$eventIDs{4625} = 'An account failed to log on';
	$eventIDs{4634} = 'An account was logged off';
	$eventIDs{4647} = 'User initiated logoff';
	$eventIDs{4778} = 'A session was reconnected to a Window Station';
	$eventIDs{4779} = 'A session was disconnected from a Window Station';
	$eventIDs{9009} = 'The Desktop Window Manager has exited with code (<X>).'; # Must find Instance ID for this one
	my %eventTypes = (
		0		=> 'Success',
		1		=> 'Error',
		2		=> 'Warning',
		4		=> 'Information',
		8 	=> 'Audit Success',
		16	=> 'Audit Failure',
	);
	my %logonFailures = (
		'0xc000005E' => 'There are currently no logon servers available to service the logon request.',
		'0xc0000064' => 'user name does not exist',
		'0xc000006A' => 'user name is correct but the password is wrong',
		'0xc000006D' => 'This is either due to a bad username or authentication information',
		'0xc000006E' => 'Unknown user name or bad password.',
		'0xc000006F' => 'user tried to logon outside his day of week or time of day restrictions',
		'0xc0000070' => 'workstation restriction, or Authentication Policy Silo violation (look for event ID 4820 on domain controller)',
		'0xc0000071' => 'expired password',
		'0xc0000072' => 'account is currently disabled',
		'0xc00000DC' => 'Indicates the Sam Server was in the wrong state to perform the desired operation.',
		'0xc0000133' => 'clocks between DC and other computer too far out of sync',
		'0xc000015b' => 'The user has not been granted the requested logon type (aka logon right) at this machine',
		'0xc000018C' => 'The logon request failed because the trust relationship between the primary domain and the trusted domain failed.',
		'0xc0000192' => 'An attempt was made to logon, but the netlogon service was not started.',
		'0xc0000193' => 'account expiration',
		'0xc0000224' => 'user is required to change password at next logon',
		'0xc0000225' => 'evidently a bug in Windows and not a risk',
		'0xc0000234' => 'user is currently locked out',
		'0xc0000413' => 'Logon Failure: The machine you are logging onto is protected by an authentication firewall. The specified account is not allowed to authenticate to the machine.',
	);
	my %disconnectCode = (
		1 => 'The disconnection was initiated by an administrative tool on the server in another session.',
		2 => 'The disconnection was due to a forced logoff initiated by an administrative tool on the server in another session.',
		3 => 'The idle session limit timer on the server has elapsed.',
		4 => 'The active session limit timer on the server has elapsed.',
		5 => 'Another user connected to the server, forcing the disconnection of the current connection.',
		6 => 'The server ran out of available memory resources.',
		7 => 'The server denied the connection.',
		9 => 'The user cannot connect to the server due to insufficient access privileges.',
		10 => 'The server does not accept saved user credentials and requires that the user enter their credentials for each connection.',
		11 => 'The disconnection was initiated by the user disconnecting his or her session on the server or by an administrative tool on the server.',
		12 => 'The disconnection was initiated by the user logging off his or her session on the server.',
	);
	# Parse the log
	my $nbrActivities = 0;
	if (my $parser = Win32::EventLog->new($file)) {
		print "Parsing $file...\n";
		my $lastUT;
		my $nbrRecs;
		my $base;
		my $x = 0;
		$parser->GetNumber($nbrRecs);
		$parser->GetOldest($base);
		while ($x < $nbrRecs) {
			my $hashRef;
			my $ip;
			if ($parser->Read(EVENTLOG_FORWARDS_READ|EVENTLOG_SEEK_READ,$base+$x,$hashRef)) {
				Win32::EventLog::GetMessageText($hashRef);
				print "First entry date: " . &date($hashRef->{TimeGenerated}) . "\n" if !$x;
				$lastUT = $hashRef->{TimeGenerated};
				if ((!$DATE_START or ($DATE_START and $lastUT >= $DATE_START)) and (!$DATE_END or ($DATE_END and $lastUT < $DATE_END))) {
					if (($elog == 1 and ($hashRef->{EventID} =~ /^(?:4624|4625|4634|4647|4778|4779)$/)) or
							($elog == 2 and ($hashRef->{EventID} =~ /^(?:21|22|23|24|25|39|40)$/)) or
							($elog == 3 and ($hashRef->{EventID} =~ /^(?:261|1149)$/)) or
							($elog == 4 and ($hashRef->{EventID} =~ /^(?:-1073086408)$/))) {
						$hashRef->{EventID} =~ s/-1073086408/56/; # Instance ID to Event ID (System.evtx)
						if (($hashRef->{Strings} or $hashRef->{Message} or ($hashRef->{EventID} and $hashRef->{EventID} == 261)) and
								$eventIDs{$hashRef->{EventID}}) {
							if ($TYPE != 5) { # Remove everything that don't contain an IP address
								if (($hashRef->{Strings} and $hashRef->{Strings} =~ /(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/) or
										($hashRef->{Message} and $hashRef->{Message} =~ /(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/)) {
									$ip = $1;
									if ($ip =~ /(?:^127\.|^10\.|^172\.1[6-9]\.|^172\.2[0-9]\.|^172\.3[0-1]\.|^192\.168\.)/ and ($TYPE == 1 or $TYPE == 3)) {
										$x++;
										next;
									}
								} else { $x++; next; }
							}
							my $eventID;
							my $eventInd = $hashRef->{TimeGenerated}.'-'.$hashRef->{RecordNumber};
							foreach my $name (%{$hashRef}) {
								if ($name and defined($hashRef->{$name}) and $hashRef->{$name}) {
									my $value;
									if 		($name eq 'EventID'								) { $value	 = $hashRef->{$name} . ' - ' . $eventIDs{$hashRef->{$name}};
																															$eventID = $hashRef->{$name}; }
									elsif ($name eq 'EventType'							) { $value	 = $hashRef->{$name} . ' - ' . $eventTypes{$hashRef->{$name}}; }
									elsif ($name eq 'Strings' and $TYPE <= 2) { $value	 = $ip; }
									elsif ($name eq 'Strings' and $elog			) { $value	 = &formatEventStrings($hashRef->{$name}, $eventID, \%logonFailures,
																																														 \%disconnectCode); }
									else																			{ $value = $hashRef->{$name}; }
									$$refEvents{$eventInd}{$name} = $value if $value;
								}
							}
							$$refEvents{$eventInd}{Strings} = $$refEvents{$eventInd}{Message} if $hashRef->{EventID} == 56 and $TYPE > 2;
							$nbrActivities++;
						}
					}
				}
			}
			$x++;
		}
		print "Last entry date: " . &date($lastUT) . "\n";
		print "Number of entries: $x\n\n";
		return($nbrActivities);
	}
	
}   #--- End parseEventLog

#--------------------------#
sub formatEventStrings
#--------------------------#
{
	# Local variables
	my ($eventStr, $eventID, $refLogonFailures, $refDisconnectCode) = @_;
	my @strings = split(/\0+/, $eventStr);
	my $formatedString;
	if ($eventID < 30) {
		$formatedString .= "User: $strings[0]\r\n";
		$formatedString .= "Session ID: $strings[1]\r\n";
		$formatedString .= "Source Network Address:	$strings[2]" if $eventID != 23;
	} elsif ($eventID == 39) {
		$formatedString .= "Session ID: $strings[0]\r\n";
		$formatedString .= "Session ID: $strings[1]\r\n";
	} elsif ($eventID == 40) {
		$formatedString .= "Session ID: $strings[0]\r\n";
		if ($strings[1] and $$refDisconnectCode{$strings[1]}) {
			$formatedString .= "Reason: $$refDisconnectCode{$strings[1]}\r\n";
		} elsif ($strings[1]) {
			$formatedString .= "Reason code: $strings[1]\r\n";
		} else {
			$formatedString .= "Reason code: 0\r\n";
		}
	} elsif ($eventID == 1149) {
		$formatedString .= "Remote Desktop Services: User authentication succeeded:\r\n";
		$formatedString .= "User:	$strings[0]\r\n";
		my $row = 1;
		if ($strings[3]) { $formatedString .= "Domain: $strings[$row]\r\n"; $row++; }
		$formatedString .= "Source Network Address:	$strings[$row]\r\n";
	} elsif ($eventID == 4648) {
		$formatedString .= "Subject:\r\n";
		$formatedString .= "Security ID: $strings[0]\r\n";
		$formatedString .= "Account Name: $strings[1]\r\n";
		$formatedString .= "Account Domain:	$strings[2]\r\n";
		$formatedString .= "Logon ID: $strings[3]\r\n";
		$formatedString .= "Logon GUID:	$strings[4]\r\n\r\n";
		$formatedString .= "Account Whose Credentials Were Used:\r\n";
		$formatedString .= "Account Name: $strings[5]\r\n";
		$formatedString .= "Account Domain:	$strings[6]\r\n";
		$formatedString .= "Logon GUID:	$strings[7]\r\n\r\n";
		$formatedString .= "Target Server:\r\n";
		$formatedString .= "Target Server Name:	$strings[8]\r\n";
		$formatedString .= "Additional Information:	$strings[9]\r\n\r\n";
		$formatedString .= "Process Information:\r\n";
		$formatedString .= "Process ID:	$strings[10]\r\n";
		$formatedString .= "Process Name: $strings[11]\r\n\r\n";
		$formatedString .= "Network Information:\r\n";
		$formatedString .= "Network Address: $strings[12]\r\n";
		$formatedString .= "Port: $strings[13]";
	} elsif ($eventID == 4624) {
		$formatedString .= "Subject:\r\n";
		$formatedString .= "Security ID: $strings[0]\r\n";
		$formatedString .= "Account Name: $strings[1]\r\n";
		$formatedString .= "Account Domain:	$strings[2]\r\n";
		$formatedString .= "Logon ID: $strings[3]\r\n";
		$formatedString .= "Logon GUID:	$strings[4]\r\n\r\n";
		$formatedString .= "Account Whose Credentials Were Used:\r\n";
		$formatedString .= "Account Name: $strings[5]\r\n";
		$formatedString .= "Account Domain:	$strings[6]\r\n";
		$formatedString .= "Logon ID: $strings[7]\r\n\r\n";
		$formatedString .= "Logon Type: $strings[8]\r\n";
		$formatedString .= "Detailed Authentication Information:\r\n";
		$formatedString .= "Logon Process: $strings[9]\r\n";
		$formatedString .= "Authentication Package: $strings[10]\r\n";
		my $row = 11;
		if ($strings[$row] !~ /^{/) {
			$formatedString .= "Account Domain: $strings[$row]\r\n"; $row++;
		}
		$formatedString .= "Logon GUID: $strings[$row]\r\n"; $row++;
		$formatedString .= "Transited Services: $strings[$row]\r\n"; $row++;
		$formatedString .= "Package Name (NTLM only): $strings[$row]\r\n"; $row++;
		$formatedString .= "Key Length: $strings[$row]\r\n\r\n"; $row++;
		$formatedString .= "Process Information:\r\n";
		$formatedString .= "Process ID: $strings[$row]\r\n"; $row++;
		$formatedString .= "Process Name: $strings[$row]\r\n\r\n"; $row++;
		$formatedString .= "Network Information:\r\n";
		$formatedString .= "Network Address: $strings[$row]\r\n"; $row++;
		if ($strings[$row]) { $formatedString .= "Port: $strings[$row]"; }
		else								{ $formatedString .= "Port: -";							 }
	} elsif ($eventID == 4625) {
		$formatedString .= "Subject:\r\n";
		$formatedString .= "Security ID: $strings[0]\r\n";
		$formatedString .= "Account Name: $strings[1]\r\n";
		$formatedString .= "Account Domain:	$strings[2]\r\n";
		$formatedString .= "Logon ID: $strings[3]\r\n";
		$formatedString .= "Account For Which Logon Failed:\r\n";
		$formatedString .= "Security ID: $strings[4]\r\n";
		$formatedString .= "Account Name: $strings[5]\r\n";
		$formatedString .= "Account Domain:	\r\n";
		my $failReason;
		if (exists($$refLogonFailures{$strings[6]})) {
			$failReason = $$refLogonFailures{$strings[6]};
		} elsif (exists($$refLogonFailures{$strings[8]})) {
			$failReason = $$refLogonFailures{$strings[8]};
		}
		$formatedString .= "Failure Information:\r\n";
		$formatedString .= "Failure Reason: $failReason\r\n" if $failReason;
		$formatedString .= "Status: $strings[6]\r\n";
		$formatedString .= "Sub Status:	$strings[8]\r\n";
		$formatedString .= "Logon Type: $strings[9]\r\n";
		$formatedString .= "Logon Process: $strings[10]\r\n";
		$formatedString .= "Authentication Package: $strings[11]\r\n";
		my $row = 12;
		if ($strings[$row] !~ /^-/) {
			$formatedString .= "Workstation Name: $strings[$row]\r\n"; $row++;
		}
		$formatedString .= "Transited Services: $strings[$row]\r\n"; $row++;
		$formatedString .= "Package Name (NTLM only): $strings[$row]\r\n"; $row++;
		$formatedString .= "Key Length: $strings[$row]\r\n"; $row++;
		$formatedString .= "Caller Process ID: $strings[$row]\r\n"; $row++;
		$formatedString .= "Caller Process Name: $strings[$row]\r\n"; $row++;
		$formatedString .= "Source Network Address: $strings[$row]\r\n"; $row++;
		$formatedString .= "Source Port: $strings[$row]";
	} elsif ($eventID == 4647) {
		$formatedString .= "Security ID:	$strings[0]\r\n";
		$formatedString .= "Account Name: $strings[1]\r\n";
		$formatedString .= "Account Domain:	$strings[2]\r\n";
		$formatedString .= "Logon ID: $strings[3]";
	} elsif ($eventID == 4634) {
		$formatedString .= "Security ID:	$strings[0]\r\n";
		$formatedString .= "Account Name: $strings[1]\r\n";
		$formatedString .= "Account Domain:	$strings[2]\r\n";
		$formatedString .= "Logon ID: $strings[3]\r\n";
		$formatedString .= "Logon Type: $strings[4]";
	} else {
		my $i = 0;
		foreach (@strings) { chomp($strings[$i]); $formatedString .= "$strings[$i]\r\n"; $i++; }
	}
	return($formatedString);
	
}   #--- End formatEventStrings

#--------------------------#
sub parseOldEventLog
#--------------------------#
{
	# Local variables
	my ($file, $refEvents) = @_;
	my %eventIDs = (
		# File is security.evt or SecEvent.Evt (old format)
		528 => 'Successful Logon',
		529 => 'Logon Failure - Unknown user name or bad password',
		530 => 'Logon Failure - Account logon time restriction violation',
		531 => 'Logon Failure - Account currently disabled',
		532 => 'Logon Failure - The specified user account has expired',
		533 => 'Logon Failure - User not allowed to logon at this computer',
		534 => 'Logon Failure - The user has not been granted the requested logon type at this machine',
		535 => "Logon Failure - The specified account's password has expired",
		536 => 'Logon Failure - The NetLogon component is not active',
		537 => 'Logon failure - The logon attempt failed for other reasons.',
		538 => 'User Logoff',
		539 => 'Logon Failure - Account locked out',
		540 => 'Successful Network Logon',
		552 => 'Logon attempt using explicit credentials',
		682 => 'Session reconnected to winstation',
		683 => 'Session disconnected from winstation',
	);
	my %eventTypes = (
		0		=> 'Success',
		1		=> 'Error',
		2		=> 'Warning',
		4		=> 'Information',
		8 	=> 'Audit Success',
		16	=> 'Audit Failure',
	);
	my %ips;
	# Parse the log
	my $nbrActivities = 0;
	print "Opening file... (this process may take a while)\n";
	if (my $parser = Parse::EventLog->new($file)) {
		print "Parsing $file...\n";
		my $lastUT;
		my $base;
		my $x;
		my $first			= 0;
		my %allEvents = $parser->getAll();
		my $nbrRecs		= scalar(keys %allEvents);
		for (my $k = 0; $k < 2; $k++) { # 2 pass
			$x = 0;
			foreach my $ind (sort keys %allEvents) {
				my $hashRef;
				my $ip;
				my $eventTimeUT;
				if ($allEvents{$ind}{TimeGenerated} =~ /\d{10}/) {
					$eventTimeUT = $allEvents{$ind}{TimeGenerated};
					if (!$first) { print "First entry date: " . &date($eventTimeUT) . "\n"; $first++; }
					$lastUT = $eventTimeUT;
					if ((!$DATE_START or ($DATE_START and $eventTimeUT >= $DATE_START)) and (!$DATE_END or ($DATE_END and $eventTimeUT < $DATE_END))) {
						my $eventString;
						if ($allEvents{$ind}{Strings} and $eventIDs{$allEvents{$ind}{EventID}}) { # Only defined events
							$eventString = join("\n", @{$allEvents{$ind}{Strings}});
							if (($eventString =~ /RDP-/ or $eventString =~ /\n10\n/) and $eventString =~ /(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/) {
								# Parsing string for event id 528 (and probably the whole range 528 to 540) is buggy
								# You're not gonna get all events, because strings are not correctly parsed
								$ip = $1;
								# Remove everything that don't contain an external IP address
								if ($ip =~ /(?:^127\.|^10\.|^172\.1[6-9]\.|^172\.2[0-9]\.|^172\.3[0-1]\.|^192\.168\.)/ and ($TYPE == 1 or $TYPE == 3)) { $x++; next; }
								$ips{$ip} = 1 if !exists($ips{$ip});
							} elsif ($eventString =~ /(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/) { # Not specific RDP event, but related to same IPs
								$ip = $1;
								if (!exists($ips{$ip})) { $x++; next; }
							} else { $x++; next; }
							my $eventID;
							my $eventInd = $eventTimeUT.'-'.$allEvents{$ind}{RecordNumber};
							if (!defined($$refEvents{$eventInd})) {
								foreach my $name (%{$allEvents{$ind}}) {
									if ($name and defined($allEvents{$ind}{$name}) and $allEvents{$ind}{$name}) {
										my $value;
										if 		($name eq 'EventID'								) { $value 	 = $allEvents{$ind}{$name} . ' - ' . $eventIDs{$allEvents{$ind}{$name}};
																																$eventID = $allEvents{$ind}{$name}; }
										elsif ($name eq 'EventType'							) { $value = GetEventType($allEvents{$ind}{$name}) }
										elsif ($name eq 'Strings' and $TYPE <= 2) { $value = $ip; }
										elsif ($name eq 'Strings'								) { $value = &formatOldEventStrings($eventString, $eventID); }
										else																			{ $value = $allEvents{$ind}{$name}; }
										$$refEvents{$eventInd}{$name} = $value if $value;
									}
								}
								$nbrActivities++;
							}
						}
					}
				}
				$x++;
			}
			if (!$k) {
				print "Last entry date: " . &date($lastUT) . "\n";
				print "Number of entries: $x\n\n";
			}
		}
		return($nbrActivities);
	}
	
}   #--- End parseOldEventLog

#--------------------------#
sub formatOldEventStrings
#--------------------------#
{
	# Local variables
	my ($eventStr, $eventID) = @_;
	my @strings = split(/\n/, $eventStr);
	my $formatedString;
	if (($eventID >= 528 and $eventID <= 536) or $eventID <= 539 or $eventID <= 540) {
		$formatedString .= "User Name: $strings[0]\r\n";
		$formatedString .= "Domain: $strings[1]\r\n";
		$formatedString .= "Logon ID:	$strings[2]\r\n";
		$formatedString .= "Logon Type: $strings[3]\r\n";
		$formatedString .= "Logon Process:	$strings[4]\r\n";
		$formatedString .= "Authentication Package: $strings[5]\r\n";
		$formatedString .= "Workstation Name:	$strings[6]\r\n";
		$formatedString .= "Logon GUID:	$strings[7]";
		if ($strings[8]) {
			$formatedString .= "\r\n";
			$formatedString .= "Caller User Name:	$strings[8]\r\n";
			$formatedString .= "Caller Domain:	$strings[9]\r\n";
			$formatedString .= "Caller Logon ID: $strings[10]\r\n";
			$formatedString .= "Caller Process ID:	$strings[11]\r\n";
			$formatedString .= "Transited Services: $strings[12]\r\n";
			$formatedString .= "Source Network Address: $strings[13]\r\n";
			$formatedString .= "Source Port: $strings[14]";
		}
	} elsif ($eventID == 537) {
		$formatedString .= "User Name: $strings[0]\r\n";
		$formatedString .= "Domain: $strings[1]\r\n";
		$formatedString .= "Logon Type: $strings[2]\r\n";
		$formatedString .= "Logon Process:	$strings[3]\r\n";
		$formatedString .= "Authentication Package: $strings[4]\r\n";
		$formatedString .= "Workstation Name:	$strings[5]\r\n";
		$formatedString .= "Status code: $strings[6]\r\n";
		$formatedString .= "Substatus code:	$strings[7]\r\n";
		$formatedString .= "Caller User Name:	$strings[8]\r\n";
		$formatedString .= "Caller Domain:	$strings[9]\r\n";
		$formatedString .= "Caller Logon ID: $strings[10]\r\n";
		$formatedString .= "Caller Process ID:	$strings[11]\r\n";
		$formatedString .= "Transited Services: $strings[12]\r\n";
		$formatedString .= "Source Network Address: $strings[13]\r\n";
		$formatedString .= "Source Port: $strings[14]";
	} elsif ($eventID == 538) {
		$formatedString .= "User Name: $strings[0]\r\n";
		$formatedString .= "Domain: $strings[1]\r\n";
		$formatedString .= "Logon ID:	$strings[2]\r\n";
		$formatedString .= "Logon Type: $strings[3]";
	} elsif ($eventID == 552) {
		my $row = 0;
		$row++ if length($strings[0]) < 3;
		$formatedString .= "Logged on user:\r\n";
		$formatedString .= "User Name: $strings[$row]\r\n"; $row++;
		$formatedString .= "Domain: $strings[$row]\r\n"; $row++;
		$formatedString .= "Logon ID:	$strings[$row]\r\n"; $row++;
		$formatedString .= "Logon GUID:	$strings[$row]\r\n\r\n"; $row++;
		$formatedString .= "User whose credentials were used:\r\n";
		$formatedString .= "Target User Name: $strings[$row]\r\n"; $row++;
		$formatedString .= "Target Domain:	$strings[$row]\r\n"; $row++;
		$formatedString .= "Target Logon GUID: $strings[$row]\r\n\r\n"; $row++;
		$formatedString .= "Target Server Name: $strings[$row]\r\n"; $row++;
		$formatedString .= "Target Server Info: $strings[$row]\r\n"; $row++;
		$formatedString .= "Caller Process ID: $strings[$row]\r\n"; $row++;
		$formatedString .= "Source Network Address: $strings[$row]\r\n"; $row++;
		if ($strings[$row]) { $formatedString .= "Source Port: $strings[$row]";	}
		else								{ $formatedString .= "Source Port: -";							}
	} elsif ($eventID == 682 or $eventID == 683) {
		$formatedString .= "User Name: $strings[0]\r\n";
		$formatedString .= "Domain:	$strings[1]\r\n";
		$formatedString .= "Logon ID: $strings[2]\r\n";
		$formatedString .= "Session Name: $strings[3]\r\n";
		$formatedString .= "Client Name:	$strings[4]\r\n";
		$formatedString .= "Client Address: $strings[5]";
	} else {
		my $i = 0;
		foreach (@strings) { chomp($strings[$i]); $formatedString .= "$strings[$i]\r\n"; $i++; }
	}
	return($formatedString);
	
}   #--- End formatOldEventStrings

#--------------------------#
sub reportXLSX
#--------------------------#
{
	# Local variables
	my ($report, $refEvents) = @_;
  # Create an XLSX workbook with a single sheet
  if (my $excel = Excel::Writer::XLSX->new($report)) {
    # Set metadata
    $excel->set_properties(comments => 'Generated by RDP-Parser '.$VERSION);
    # Create a sheet
    if (my $sheet = $excel->add_worksheet('Results')) {
			# Format
			my $formatHeader	= $excel->add_format(align => 'center', size => 10, bold => 1);
			my $formatDate		= $excel->add_format(align => 'top'		, size => 10, num_format => 'yyyy-mm-dd hh:mm:ss', size => 10);
			my $formatNormal	= $excel->add_format(align => 'top'		, size => 10);
			my $formatDetails	= $excel->add_format(align => 'top'		, size => 10, text_wrap => 1);
			# Column width
      my $maxWidthCol1 = length('TimeGenerated')+5;
      my $maxWidthCol2 = length('Timewritten')	+5;
      my $maxWidthCol3 = length('Computer')			+5;
      my $maxWidthCol4 = length('Source')				+5;
      my $maxWidthCol5 = length('RecordNumber')	+5;
      my $maxWidthCol6 = length('Category')			+5;
      my $maxWidthCol7 = length('EventID')			+5;
      my $maxWidthCol8 = length('EventType')		+5;
      my $maxWidthCol9 = length('Details')			+5;
			# Print header
			my $header;
			my $countCol = 0;
			my $j = 0; # Column No
			$sheet->write(0, $j, 'TimeGenerated', $formatHeader ); $j++;
			if ($TYPE > 2) { $sheet->write(0, $j, 'TimeWritten'	, $formatHeader ); $j++; }
			if ($TYPE > 2) { $sheet->write(0, $j, 'Computer'			, $formatHeader ); $j++; }
			$sheet->write(0, $j, 'Source'				, $formatHeader ); $j++;
			if ($TYPE > 2) { $sheet->write(0, $j, 'RecordNumber'	, $formatHeader ); $j++; }
			if ($TYPE > 2) { $sheet->write(0, $j, 'Category'			, $formatHeader ); $j++; }
			$sheet->write(0, $j, 'EventID'			, $formatHeader ); $j++;
			if ($TYPE > 2) { $sheet->write(0, $j, 'EventType'		, $formatHeader ); $j++; }
			$sheet->write(0, $j, 'Details'			, $formatHeader );
			# Print data
			my $i = 1;
			foreach my $ind (sort %{$refEvents}) {
				if ($ind and $$refEvents{$ind}{TimeGenerated}) {
					$j = 0;
					my $dateStr = &date($$refEvents{$ind}{TimeGenerated});
					$sheet->write_date_time($i, $j, $dateStr, $formatDate); $j++;
					$maxWidthCol1 = length($dateStr) if length($dateStr) > $maxWidthCol1;
					if ($TYPE > 2) {
						$dateStr = &date($$refEvents{$ind}{Timewritten}) if exists($$refEvents{$ind}{Timewritten});
						if ($dateStr) {
							$sheet->write_date_time($i, $j, $dateStr, $formatDate);
							$maxWidthCol2 = length($dateStr) if length($dateStr) > $maxWidthCol2;
						}
						$j++;
					}
					if ($TYPE > 2) {
						if ($$refEvents{$ind}{Computer}) {
							$sheet->write_string($i, $j, $$refEvents{$ind}{Computer}, $formatNormal);
							$maxWidthCol3 = length($$refEvents{$ind}{Computer}) if length($$refEvents{$ind}{Computer}) > $maxWidthCol3;
						}
						$j++;
					}
					if ($$refEvents{$ind}{Source}) {
						$sheet->write_string($i, $j, $$refEvents{$ind}{Source}, $formatNormal);
						$maxWidthCol4 = length($$refEvents{$ind}{Source}) if length($$refEvents{$ind}{Source}) > $maxWidthCol4;
					}
					$j++;
					if ($TYPE > 2) {
						if ($$refEvents{$ind}{RecordNumber}) {
							$sheet->write_string($i, $j, $$refEvents{$ind}{RecordNumber}, $formatNormal);
							$maxWidthCol5 = length($$refEvents{$ind}{RecordNumber}) if length($$refEvents{$ind}{RecordNumber}) > $maxWidthCol5;
						}
						$j++;
					}
					if ($TYPE > 2) {
						if ($$refEvents{$ind}{Category}) {
							$sheet->write_string($i, $j, $$refEvents{$ind}{Category}, $formatNormal) if $$refEvents{$ind}{Category};
							$maxWidthCol6 = length($$refEvents{$ind}{Category}) if length($$refEvents{$ind}{Category}) > $maxWidthCol6;
						}
						$j++;
					}
					if ($$refEvents{$ind}{EventID}) {
						$sheet->write_string($i, $j, $$refEvents{$ind}{EventID}, $formatNormal); 
						$maxWidthCol7 = length($$refEvents{$ind}{EventID}) if length($$refEvents{$ind}{EventID}) > $maxWidthCol7;
					}
					$j++;
					if ($TYPE > 2) {
						if ($$refEvents{$ind}{EventType}) {
							$sheet->write_string($i, $j, $$refEvents{$ind}{EventType}, $formatNormal);
							$maxWidthCol8 = length($$refEvents{$ind}{EventType}) if length($$refEvents{$ind}{EventType}) > $maxWidthCol8;
						}
						$j++;
					}
					if ($$refEvents{$ind}{Strings}) {
						chop($$refEvents{$ind}{Strings}) while $$refEvents{$ind}{Strings} =~ /[\r\n]$/;
						if ($DATA_STR) {
							$$refEvents{$ind}{Strings} =~ s/\r\n/\|/g;
							$sheet->write_string($i, $j, $$refEvents{$ind}{Strings}, $formatNormal);
						} else {
							$sheet->write_string($i, $j, $$refEvents{$ind}{Strings}, $formatDetails);
							$maxWidthCol9 = length($$refEvents{$ind}{Strings}) if length($$refEvents{$ind}{Strings}) > $maxWidthCol9;
						}
					}
					$i++;
				}
			}
      # Ajust column sizes
      $j = 0;
      $sheet->set_column($j, $j, $maxWidthCol1); $j++;
      if ($TYPE > 2	) { $sheet->set_column($j, $j, $maxWidthCol2); $j++; }
			if ($TYPE > 2	) { $sheet->set_column($j, $j, $maxWidthCol3); $j++; }
      $sheet->set_column($j, $j, $maxWidthCol4); $j++;
			if ($TYPE > 2	) { $sheet->set_column($j, $j, $maxWidthCol5); $j++; }
			if ($TYPE > 2	) { $sheet->set_column($j, $j, $maxWidthCol6); $j++; }
      $sheet->set_column($j, $j, $maxWidthCol7); $j++;
			if ($TYPE > 2	) { $sheet->set_column($j, $j, $maxWidthCol8); $j++; }
      if ($TYPE <= 2) { $sheet->set_column($j, $j, $maxWidthCol9); }
			else						{ $sheet->set_column($j, $j, 80);						 }
      # Set autofilter
      $sheet->autofilter(0, 0, 0, $j);
      $sheet->freeze_panes(1, 0);
		}
		$excel->close(); # Close the file
		return($report);
	}
	
}  #--- End reportXLSX

#--------------------------#
sub reportHTML
#--------------------------#
{
	# Local variables
	my ($report, $refEvents) = @_;
  # Create an XLSX workbook with a single sheet
  if (open(my $fhReport, '>', $report)) {
    flock($fhReport, 2);
		my $timeStr = &date(time);
    print $fhReport "<!DOCTYPE html>\n";
    print $fhReport "<html>\n<head>\n<title>RDP-Parser report $timeStr</title>\n";
    print $fhReport "<meta name=\"generator\" content=\"RDP-Parser $VERSION\">\n";
    print $fhReport "<style>\n";
    print $fhReport "table, td { border-collapse: collapse; border: 1px solid black; padding: 5px; }\n";
    print $fhReport "td { font-size:11pt; vertical-align: top; white-space: nowrap; }\n";
    print $fhReport ".header { text-align: center; font-weight: bold }\n";
    print $fhReport "</style>\n";
    print $fhReport "</head>\n";
    print $fhReport "<body style=\"font-family: Calibri, Verdana, Arial;\">\n";
    print $fhReport "<table align=\"center\">\n";
    # File Header
		print $fhReport "<tr>\n";
		print $fhReport "<td class=\"header\">TimeGenerated</td>\n";
		print $fhReport "<td class=\"header\">Timewritten</td>\n"		if $TYPE > 2;
		print $fhReport "<td class=\"header\">Computer</td>\n"			if $TYPE > 2;
		print $fhReport "<td class=\"header\">Source</td>\n";
		print $fhReport "<td class=\"header\">RecordNumber</td>\n"	if $TYPE > 2;
		print $fhReport "<td class=\"header\">Category</td>\n"			if $TYPE > 2;
		print $fhReport "<td class=\"header\">EventID</td>\n";
		print $fhReport "<td class=\"header\">EventType</td>\n"			if $TYPE > 2;
		print $fhReport "<td class=\"header\">Details</td>\n";
		print $fhReport "</tr>\n";
		# Print data
		my $i = 1;
		foreach my $ind (sort %{$refEvents}) {
			if ($ind and $$refEvents{$ind}{TimeGenerated}) {
				print $fhReport "<tr>\n";
				my $dateStr = &date($$refEvents{$ind}{TimeGenerated});
				print $fhReport "<td>$dateStr</td>\n";
				if ($TYPE > 2) {
					if ($$refEvents{$ind}{Timewritten}) {
						$dateStr = &date($$refEvents{$ind}{Timewritten});
						print $fhReport "<td>$dateStr</td>\n";
					} else { print $fhReport "<td></td>\n"; }
				}
				print $fhReport "<td>$$refEvents{$ind}{Computer}</td>\n"			if $TYPE > 2;
				print $fhReport "<td>$$refEvents{$ind}{Source}</td>\n";
				print $fhReport "<td>$$refEvents{$ind}{RecordNumber}</td>\n"	if $TYPE > 2;
				if ($TYPE > 2) {
					if ($$refEvents{$ind}{Category}) {
						print $fhReport "<td>$$refEvents{$ind}{Category}</td>\n";
					} else { print $fhReport "<td></td>\n"; }
				}
				print $fhReport "<td>$$refEvents{$ind}{EventID}</td>\n";
				print $fhReport "<td>$$refEvents{$ind}{EventType}</td>\n"			if $TYPE > 2;
				if ($TYPE <= 2) { print $fhReport "<td>$$refEvents{$ind}{Strings}</td>\n"; }
				else {
					chop($$refEvents{$ind}{Strings}) while $$refEvents{$ind}{Strings} =~ /[\r\n]$/;
					my $value = $$refEvents{$ind}{Strings};
					if ($DATA_STR)	{ $value =~ s/\r\n/\|/g; 	 }
					else						{ $value =~ s/\r\n/<br>/g; }
					print $fhReport "<td>$value</td>\n";
				}
			}
		}
		close($report);
		return($report);
	}
	
}  #--- End reportHTML

#--------------------------#
sub reportTXT
#--------------------------#
{
	# Local variables
	my ($report, $refEvents) = @_;
  # Create an XLSX workbook with a single sheet
  if (open(my $fhReport, '>', $report)) {
    flock($fhReport, 2);
    # File Header
		print $fhReport "TimeGenerated\t";
		print $fhReport "Timewritten\t"		if $TYPE > 2;
		print $fhReport "Computer\t"			if $TYPE > 2;
		print $fhReport "Source\t";
		print $fhReport "RecordNumber\t"	if $TYPE > 2;
		print $fhReport "Category\t"			if $TYPE > 2;
		print $fhReport "EventID\t";
		print $fhReport "EventType\t"			if $TYPE > 2;
		print $fhReport "Details\n";
		# Print data
		my $i = 1;
		foreach my $ind (sort %{$refEvents}) {
			if ($ind and $$refEvents{$ind}{TimeGenerated}) {
				my $dateStr = &date($$refEvents{$ind}{TimeGenerated});
				print $fhReport "$dateStr\t";
				if ($TYPE > 2) {
					if ($$refEvents{$ind}{Timewritten}) {
						$dateStr = &date($$refEvents{$ind}{Timewritten});
						print $fhReport "$dateStr";
					}
					print $fhReport "\t";
				}
				print $fhReport "$$refEvents{$ind}{Computer}\t"			if $TYPE > 2;
				print $fhReport "$$refEvents{$ind}{Source}\t";
				print $fhReport "$$refEvents{$ind}{RecordNumber}\t" if $TYPE > 2;
				if ($TYPE > 2) {
					print $fhReport "$$refEvents{$ind}{Category}" if $$refEvents{$ind}{Category};
					print $fhReport "\t";
				}
				print $fhReport "$$refEvents{$ind}{EventID}\t";
				print $fhReport "$$refEvents{$ind}{EventType}\t" if $TYPE > 2;
				if ($TYPE <= 2) { print $fhReport "$$refEvents{$ind}{Strings}\n"; }
				else {
					chop($$refEvents{$ind}{Strings}) while $$refEvents{$ind}{Strings} =~ /[\r\n]$/;
					if ($DATA_STR)	{
						$$refEvents{$ind}{Strings} =~ s/\r\n/\|/g;
						$$refEvents{$ind}{Strings} =~ s/[\r\n]/\|/g;
						print $fhReport "$$refEvents{$ind}{Strings}\n";
					} else { print $fhReport "\"$$refEvents{$ind}{Strings}\"\n"; }
				}
			}
		}
		close($report);
		return($report);
	}
	
}  #--- End reportTXT

#--------------------------#
sub date
#--------------------------#
{
  # Conversion
  my ($time) = @_;
	if ($time) {
		my ($s,$min,$hr,$j,$m,$an,$jour_s,$ha,$isDST) = localtime($time);
		return(sprintf("%04d\-%02d\-%02d %02d:%02d:%02d", $an+1900, $m+1, $j, $hr, $min, $s));
	}

}  #--- End date
