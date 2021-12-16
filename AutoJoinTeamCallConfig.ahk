;///////////////////////////
; Configuration Tool to setup Automatic Conference Call dialer.
;    For details, run the script & check the help using desktray menu
; https://www.robvanderwoude.com/schtasks.php#Create
;///////////////////////////

#SingleInstance force

; Read Default settings from ini file
IniRead, adBridgeLink ,%A_ScriptDir%/AutoJoinTeamCall.ini, Settings, BridgeLink, http://
IniRead, adAudioFile  ,%A_ScriptDir%/AutoJoinTeamCall.ini, Settings, AudioFile, %A_ScriptDir%\HelloIamXfromYteam.wma

;Delete standard menu, set custom menu items
Menu, tray, NoStandard
Menu, tray, add, &Help, MenuHelp
Menu, tray, add, E&xit, ButtonCancel

;Get current local time, parse it, store lt in variables for the GUI Timer
FormatTime, CDate,, yyyy-M-d-H-m
Loop, parse, CDate, -,
{
	if A_Index = 1
		CYear=%A_LoopField%
	else if A_Index = 2
		CMonth=%A_LoopField%
	else if A_Index = 3
		CDay=%A_LoopField%
	else if A_Index = 4
		CHour=%A_LoopField%
	else
		CMin=%A_Loopfield%
}

;Start GUI with default values
Gui, Add, GroupBox, w330 h120, Timer Settings                      ; Use curent date time as default
Gui, Add, Text, ym+40 xm+10, Year | 
Gui, Add, Text, ym+40 x+5, Month | 
Gui, Add, Text, ym+40 x+5, Day
Gui, Add, Edit, ym+40 xm+150 w65
Gui, Add, UpDown, Wrap 0x80 vGUIYear Range2020-2050, %CYear%
Gui, Add, Edit, ym+40 w35 x+1
Gui, Add, UpDown, Wrap vGUIMonth Range1-12, %CMonth%
Gui, Add, Edit, ym+40 w35 x+1
Gui, Add, UpDown, Wrap vGUIDay Range1-31, %CDay%

Gui, Add, Text, ym+70 xm+10, Hour | 
Gui, Add, Text, ym+70 x+5, Minute
Gui, Add, Edit, ym+70 w35 xm+150
Gui, Add, UpDown, Wrap vGUIHour Range0-23, %CHour%
Gui, Add, Edit, ym+70 w35 x+1
Gui, Add, UpDown, Wrap vGUIMinute Range1-59, %CMin%


Gui, Add, GroupBox, x10 y160 w330 h70 , Teams/Zoom Link:             ; Use previous execution's value as default
Gui, Add, Edit, vGUIbridgeLink x16 y190 w290 h20 , %adBridgeLink%


Gui, Add, Text, ym+250 xm+30, Select audio files to play             ; Use previous execution's value as default
Gui, Add, Edit, vFileToRunEdit1 ym+280 xm+20 w236, %adAudioFile%
Gui, Add, Button, gOpenFileSelection vFileToRun1 ym+280 xm+256 w50, Browse 
GuiControl, Disable, FileToRunEdit1

Gui, Add, Button, ym+340 xm+15 w80 gScheduleCall, Schedule Call
Gui, Add, Button, ym+340 x+10 w80, Cancel
Gui, Add, Button, ym+340 x+40 w80  gMenuHelp, Help

Gui, Show, AutoSize, Auto Join Zoom or Teams Calls
Return



;File selection from the GUI button
OpenFileSelection:
Gui +OwnDialogs
FileSelectFile, RunAfterResume1, 3,, Selecting a audio file...
if RunAfterResume1 =  
    return
GuiControl,, FileToRunEdit1, %RunAfterResume1%
Return

;Cancel and Exit
ButtonCancel:
GuiClose:
ExitApp
Return

;Help menu
MenuHelp:
IfWinExist, Auto Meeting Join
{
WinActivate
}
else
{
Gui +OwnDialogs
Gui, 3:Add, Tab2, w300 h230, What is this| Troubleshoot | About 
Gui, 3:Add, Text,, Tool to schedule joining teams/zoom bridges
Gui, 3:Add, Text,w270, The tool will then connect to the bridge link at scheduled hour and play a pre-recorded audio
Gui, 3:Add, Text,w270, The tool will use Windows Schtasks to schedule the tasks and will store the configuration details in AutoJoinTeamCall.ini file.
Gui, 3:Add, Text,w270, 
Gui, 3:Add, Text,w270, This version allows scheduling of just 1 future tasks at a time.
Gui, 3:Tab, 2
Gui, 3:Add, Text, w270, Application uses windows SCHTASKS.EXE to schedule tasks. 
Gui, 3:Add, Text, w270, Use TASKSCHD.MSC to see logs if there are errors.
Gui, 3:Add, Text, w270, The ini file includes logs of past excecutions for tracking.
Gui, 3:Tab, 3 
Gui, 3:Add, Text, w270, Created by Arun Sivanandan 
Gui, 3:Add, Text,, version 1.0 ( 2021-12-16)
Gui, 3:Show, AutoSize, Auto Meeting Join
Return

3GUIClose:
Gui, 3:Destroy
Return
}
Return

ScheduleCall:
Gui, Submit  ;End of GUI
SchTime := (StrLen(GUIHour)=1 ? "0" : "") GUIHour . ":" (StrLen(GUIMinute)=1 ? "0" : "") GUIMinute                                                          ; As HH:MM
SchDate := (StrLen(GUIDay)=1 ? "0" : "") GUIDay . "/" . (StrLen(GUIMonth)=1 ? "0" : "") GUIMonth . "/" . (StrLen(GUIYear)=2 ? "20" : "") GUIYear            ; As DD/MM/YYYY

; Use windows schtasks to schedule timer
SchCmdDelTask:="SCHTASKS.exe /Delete /TN AutoJoin /f"
if FileExist(A_ScriptDir . "\AutoJoinTeamCall.exe")
  SchCmdTask:="SCHTASKS.exe /Create /SC once  /TN AutoJoin  /ST " . SchTime . " /SD "  . SchDate . " /TR """ . A_ScriptDir . "\AutoJoinTeamCall.exe"" /RU ntnet\" . A_UserName   
else
  SchCmdTask:="SCHTASKS.exe /Create /SC once  /TN AutoJoin  /ST " . SchTime . " /SD "  . SchDate . " /TR """ . A_ScriptDir . "\AutoJoinTeamCall.ahk"" /RU ntnet\" . A_UserName    

Runwait, %SchCmdDelTask%,,hide
Runwait, %SchCmdTask%,,hide UseErrorLevel
if ErrorLevel
    MsgBox "Cannot setup the scheduled tasks, check logs"

GoSub SetupTaskLogHistory
Return

SetupTaskLogHistory:     
IfInString, GUIBridgeLink, teams.microsoft         ; This is for future enhancement
    BridgeType=Teams                            
else
    BridgeType=Zoom

;Store parameters in Ini file
IniWrite, %SchDate%,         AutoJoinTeamCall.ini, Settings , ConfDate  ; Sometimes the 1st call does not update the Ini, so this is a dummy call
IniWrite, %SchDate%,         AutoJoinTeamCall.ini, Settings , ConfDate
IniWrite, %SchTime%,         AutoJoinTeamCall.ini, Settings , ConfTime
IniWrite, %GUIBridgeLink%,   AutoJoinTeamCall.ini, Settings , BridgeLink
IniWrite, %BridgeType%,      AutoJoinTeamCall.ini, Settings , BridgeType
IniWrite, %FileToRunEdit1%,  AutoJoinTeamCall.ini, Settings , AudioFile

IniWrite, %SchCmdTask%,      AutoJoinTeamCall.ini, History, SchTasksCmd%A_Now%

ExitApp
Return