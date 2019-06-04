#!/usr/bin/perl
# Perl - v: 5.16.3
#------------------------------------------------------------------------------#
# RDP-ParserLang.pl   : Strings for RDP-Parser
# Website             : http://le-tools.com/RDP-Parser.html
# SourceForge         : https://sourceforge.net/p/rdp-parser
# GitHub              : https://github.com/arioux/RDP-Parser
# Creation            : 2019-04-27
# Modified            : 2019-06-04
# Author              : Alain Rioux (admin@le-tools.com)
#
# Copyright (C) 2019  Alain Rioux (le-tools.com)
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

#------------------------------------------------------------------------------#
sub loadStr
#------------------------------------------------------------------------------#
{
  # Local variables
  my ($refSTR, $LANG_FILE) = @_;
  # Open and load string values
  open(LANG, "<:encoding(UTF-8)", $LANG_FILE);
  my @tab = <LANG>;
  close(LANG);
  # Store values  
  foreach (@tab) {
    chomp($_);
    s/[^\w\=\s\.\!\,\-\)\(\']//g;
    my ($key, $value) = split(/ = /, $_);
    $value            = encode('cp1252', $value); # Encode
    $$refSTR{$key}    = $value if $key;
  }
  
}  #--- End loadStr

#------------------------------------------------------------------------------#
sub loadDefaultStr
#------------------------------------------------------------------------------#
{
  # Local variables
  my $refSTR = shift;
  
  # Set default strings
  
  # General strings
  $$refSTR{'Error'}              = 'Error';
  $$refSTR{'Warning'}            = 'Warning';
  $$refSTR{'errorMsg'}           = 'Error messsage';
  $$refSTR{'errorConnection'}    = 'Connection error';
  $$refSTR{'errorOpening'}       = 'Error opening';
  $$refSTR{'processRunning'}     = 'A process is already running. Wait until it stops or restart the program.';
  # Main window
  $$refSTR{'Input'}           = 'Input';
  $$refSTR{'directory'}       = 'directory';
  $$refSTR{'CopyFromLive'}    = 'Copy EventLogs from current running system';
  $$refSTR{'AllEventLogs'}    = 'All Event logs';  
  $$refSTR{'Filters'}         = 'Filters';
  $$refSTR{'Events'}          = 'Events';
  $$refSTR{'EventIDs'}        = 'Event IDs';
  $$refSTR{'IPAddresses'}     = 'IP addresses';
  $$refSTR{'Select'}          = 'Select';
  $$refSTR{'NoFilter'}        = 'No filter';
  $$refSTR{'PublicIPs'}       = 'With public IPs only';
  $$refSTR{'WithIPs'}         = 'With private or public IPs';
  $$refSTR{'Dates'}           = 'Dates';
  $$refSTR{'Equal'}           = 'Equal';
  $$refSTR{'Before'}          = 'Before';
  $$refSTR{'After'}           = 'After';
  $$refSTR{'Keyword'}         = 'Keyword(s)';
  $$refSTR{'matchCase'}       = 'Match case';
  $$refSTR{'ReportOptions'}   = 'Report options';
  $$refSTR{'Path'}            = 'Path';  
  $$refSTR{'Columns'}         = 'Columns';
  $$refSTR{'All'}             = 'All';
  $$refSTR{'chReportIPsOnly'}    = 'IPs only in data';
  $$refSTR{'chReportDataInline'} = 'Data on a single line';
  $$refSTR{'Timezone'}        = 'Timezone';
  $$refSTR{'Others'}          = 'Others';
  $$refSTR{'AddStats'}        = 'Add stats to the report';
  $$refSTR{'chReportOpen'}    = 'Open report when finished';
  $$refSTR{'selDir'}          = 'Select a folder';
  $$refSTR{'openReportDir'}   = 'Open the folder in Exlorer';
  $$refSTR{'lblNotReady'}     = 'Not Ready? Click here';
  $$refSTR{'notReady'}        = 'Not ready';
  $$refSTR{'nextStep'}        = 'Next step';
  $$refSTR{'selectInput'}     = 'You must select a directory as input.';
  $$refSTR{'errNoValidDir'}   = 'You must enter a valid directory for report.';
  $$refSTR{'warnAdmin'}       = 'You must start the tool as admin to use this function.';
  $$refSTR{'selectReport'}    = 'You must select a directory for report.';
  $$refSTR{'Process'}         = 'Process';
  $$refSTR{'btnHelpTip'}      = 'See Documentation';
  $$refSTR{'Copy'}            = 'Copy';
  $$refSTR{'fileCopied'}      = 'files copied';
  $$refSTR{'Parsing'}         = 'Parsing';
  $$refSTR{'creatingReport'}  = 'Creating report';
  $$refSTR{'NoRDPFound'}      = 'No RDP activity found with selected options.';
  $$refSTR{'Stats'}           = 'Stats';
  $$refSTR{'Results'}         = 'Results';
  $$refSTR{'Computer'}        = 'Computer';
  $$refSTR{'System'}          = 'System';
  $$refSTR{'File'}            = 'File';
  $$refSTR{'FirstEntry'}      = 'First entry';
  $$refSTR{'LastEntry'}       = 'Last entry';
  $$refSTR{'NbrEntries'}      = 'Number of entries';
  # EventIDs window
  $$refSTR{'Ok'}              = 'Ok';
  $$refSTR{'CheckAll'}        = 'Check all';
  $$refSTR{'UncheckAll'}      = 'Uncheck all';
  # Config Window
  $$refSTR{'Settings'}        = 'Settings';
  $$refSTR{'general'}         = 'General';
  # General tab
  $$refSTR{'Tool'}            = 'Tool';
  $$refSTR{'Export'}          = 'Export';
  $$refSTR{'OpenUserDir'}     = 'Open user dir';
  $$refSTR{'checkUpdate'}     = 'Check Update';
  $$refSTR{'AutoUpdateTip'}   = 'Check for update at startup';
  $$refSTR{'update1'}         = 'You have the latest version installed.';
  $$refSTR{'update2'}         = 'Check for update';
  $$refSTR{'update3'}         = 'Update';
  $$refSTR{'update5'}         = 'is available. Download it';
  $$refSTR{'Functions'}       = 'Functions';
  # Logging tab
  $$refSTR{'logging'}         = 'Logging';
  $$refSTR{'chEnableLogging'} = 'Enable logging';
  $$refSTR{'OpenLog'}         = 'Open the log';
  $$refSTR{'rbUseDefaultDir'} = 'Use default folder';
  $$refSTR{'rbLoggingDir'}    = 'Use this folder';
  # About Window
  $$refSTR{'About'}           = 'About';
  $$refSTR{'Version'}         = 'Version';
  $$refSTR{'Author'}          = 'Author';
  $$refSTR{'TranslatedBy'}    = 'Translated by';
  $$refSTR{'Website'}         = 'Website';
  $$refSTR{'TranslatorName'}  = '-';
  
}  #--- End loadStrings

#------------------------------------------------------------------------------#
1;
  
  
  