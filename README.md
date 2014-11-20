Uptime
======
A simple Windows "UpTime" tool which can be used for checking the uptime of the local PC or the uptime of a remote server by network admins.  Written in vbScript, it connects to WMI to get the amount of time since the last reboot, and then calculates it into Days, hours, minutes, and seconds for display.  The script is designed to be a command line tool, and therefore if ran using the default WSCRIPT.EXE, the script will display a message warning you that it needs to be ran with CSCRIPT.EXE and provides a sample command line to set the default script executer to be CSCRIPT instead of WSCRIPT.  This script also supports the standard command line parameters, and syntax, in that it supports both dash (-) and slash (/) to identify a parameter.  Currently, the follow parameters are supported.

? – Displays Help Screen

H,HELP – Displays Help Screen

[computer/host name/IP Address] – The name of the target computer to query for Uptime.
