'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'  File   : Uptime.vbs
'  Author : Larry Frieson
'  Desc   : Gets the Windows uptime of the current machine, or another computer passed on the command line.  This can be VERY
'           useful for Network Administrators.
'  Date   : 03/14/2010
'
'  Copyright © 2010 MLinks Technologies, Inc.  All rights reserved.
'
'  Revision History: 
'    03/14/2010 17:50:43 created.
'
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
option explicit
const VERSION_INFO      = "1.0"
CheckScriptHost : CheckError("CheckScriptHost")

dim ScriptName
ScriptName = mid(wscript.ScriptName, 1, instr(wscript.ScriptName, ".") - 1)

dim bAllowWinScript
dim wsh, cmdparams, argc, fso, WMI, OperatingSystems, OS, SystemUptime, stdout
dim Days, Hours, Minutes, Seconds, SecsTotal, SecsRemaining, Computer

set wsh = CreateObject("WScript.Shell")
set fso = CreateObject("Scripting.FileSystemObject")

set cmdparams = wscript.arguments
dim bCmdSwitch

for argc = 0 to cmdparams.Count - 1
    bCmdSwitch = false
    if ((instr(cmdparams(argc), "/") = 1) or (instr(cmdparams(argc), "-") = 1)) then
        bCmdSwitch = true
        if ((lcase(mid(cmdparams(argc), 2)) = "h") > 0 OR lcase(instr(cmdparams(argc), "?")) > 0) then
            call ShowUsage(true)
            wscript.quit -1
        end if
    end if

    if bCmdSwitch = false then
        if Computer = "" then
            Computer = cmdparams(argc)
        end if
    end if
next

ShowUsage(false)

if(cmdparams.count < 1) then
    if(Computer = "") then
        Computer = wsh.ExpandEnvironmentStrings("%COMPUTERNAME%")
    end if
    if(Computer = "") then
        Computer = "."
    end if
end if

if mid(Computer, 1, 2) = "\\" then
    Computer = mid(Computer, 3)
end if
if instr(Computer, "\") > 0 then
    Computer = Replace(Computer, "\", "")
end if

wscript.stdout.Write "Connecting to \\" & Computer & "..."
Set WMI = GetObject("winmgmts:\\" & Computer & "\root\cimv2")
Set OperatingSystems = WMI.ExecQuery("Select * From Win32_PerfFormattedData_PerfOS_System")
 
For Each OS in OperatingSystems
    SystemUptime = OS.SystemUpTime
Next

' Calculate Days: save total seconds
SecsTotal     = SystemUpTime
Days          = Fix(SecsTotal / 86400)
SecsRemaining = (SecsTotal - (Days * 86400))

' Calculate Hours: save total remaining seconds
SecsTotal     = SecsRemaining
Hours         = Fix(SecsTotal / 3600)
SecsRemaining = (SecsTotal - (Hours * 3600))

' Calculate Minutes and Seconds
SecsTotal     = SecsRemaining
Minutes       = Fix(SecsTotal / 60)
Seconds       = (SecsTotal - (Minutes * 60))

wscript.stdout.Write vbCRLF & Computer & " has been online for "
wscript.stdout.WriteLine Days & " days " & Hours & " hours " & Minutes & " minutes " & Seconds & " seconds" 

sub CheckScriptHost
    dim fso
    dim strHostName
    on error resume next
    set fso = WScript.CreateObject("Scripting.FileSystemObject") : CheckError("Scripting.FileSystemObject")
    strHostName = fso.GetFileName(wscript.FullName)

    if LCase(strHostName) <> "cscript.exe" then
        wscript.echo "This script was designed to be executed from the command prompt using CSCRIPT.EXE." & vbCRLF & "For example: ""CSCRIPT.EXE """ & wscript.ScriptName & " [Options] [Parameters]""" & vbCRLF & _
                     "To set CSCRIPT.EXE as the default script host, run the following command at a command prompt, or from the Start/Run option:" & vbCRLF & "       CSCRIPT.exe /H:CSCRIPT /Nologo /S" & vbCRLF & "You can then run any VBScripts without preceding the script with cscript.exe.  " & _
                     "This makes for a MUCH better IT / Admin experience in automated batch file processing, logon script creation/execution, and many, many, more uses.  You can STILL runs script based in Windows by simply providing " & _
                     "WSCRIPT <scriptname.vbs> to execute your script in the WINDOWS Scripting Host, however I find that CScript, the Command Line Script Host is MUCH better to work with, if you plan to use vbScripting in " & _
                     "in your daily maintenance and administrative chores."
        wscript.quit 1
    end if
end sub

sub CheckError(ErrorText)
    if err = 0 then exit sub
    wscript.stdout.writeline ErrorText & ": " & err.Number & " (0x" & hex(err) & "): " & err.Description
    on error goto 0
    wscript.quit 2
end sub

sub ShowUsage(ShowHelp)
    if(ShowHelp = true) then
        wscript.stdout.WriteLine ""
    end if
    wscript.stdout.WriteLine ScriptName & " - System Uptime Tool (C) 2011 MLinks Technologies, Inc."
    wscript.stdout.WriteLine "Version " & VERSION_INFO & " by Larry Frieson" & vbCRLF
    if ShowHelp = true then
        wscript.stdout.WriteLine "  Usage: " & ScriptName & " [options] {Computername/IP Address}"
        wscript.stdout.WriteLine "  Where: [options] is one or more of the following:"
        wscript.stdout.WriteLine "         -h         Display Help"
        wscript.stdout.WriteLine "    And: {Computername} is the computer you want to check the uptime on."
        wscript.stdout.WriteLine "         NOTE: If left blank, the local computer is checked."
        wscript.stdout.WriteLine
    end if
end sub
