RDP-Parser
Description : RDP-Parser extracts RDP activities from Microsoft Windows Event Logs.
Author		: Alain Rioux (admin@le-tools.com)
WebSite		: http://le-tools.com/RDP-Parser.html
SourceForge	: https://sourceforge.net/p/RDP-Parser
GitHub		: https://github.com/arioux/RDP-Parser


Development
-----------

RDP-Parser has been developped using ActivePerl 5.16.3 with the following module installed:

- DateTime (v1.28)
- DateTime-TimeZone (v2.00)
- Excel-Writer-XLSX (v0.94)
- LWP (v6.05)
- Parse-EventLog (v0.7)
- Regexp-IPv6 (v0.03)
- Time-HiRes (v1.9726)
- Time-Local (v1.2300)
- URI (v1.71)
- Win32-API (v0.84)
- Win32-EventLog (v0.077)
- Win32-GUI (v1.11)
- Win32-Process (v0.16)
- Win32::RunAsAdmin (v0.02)
- threads (v2.21)
- threads-shared (v1.59)


ToDo
----
- Add support for Event ID 9009 in System.evtx (I didn't see an example of this Event ID it in my test)


Known problems
--------------
- For old format (evt), parsing string for event id 528 (and probably the whole range 528 to 540) is buggy, we don't get all events, because strings are not correctly parsed.



Packaging
---------

Executable has been packaged using PerlApp v.9.2.1 (ActiveState). For alternative to PerlApp, see http://www.nicholassolutions.com/tutorials/perl-PAR.htm.

Some additional modules may be required or manually added before packaging.
