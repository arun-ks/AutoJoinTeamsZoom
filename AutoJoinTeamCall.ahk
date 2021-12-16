;///////////////////////////
; Automatic Conference Call dialer.
;    The script takes configuration details form AutoJoinTeamCall.ini files & makes the call to the bridge link
;    The script's execution is scheduled using Schtasks, but can be run as a standalone too.
;///////////////////////////

#SingleInstance force
   
; Get settings from Ini file
IniRead, adConfDate   ,%A_ScriptDir%/AutoJoinTeamCall.ini, Settings, ConfDate
IniRead, adConfTime   ,%A_ScriptDir%/AutoJoinTeamCall.ini, Settings, ConfTime
IniRead, adBridgeUser ,%A_ScriptDir%/AutoJoinTeamCall.ini, Settings, BridgeLink
IniRead, adBridgeType ,%A_ScriptDir%/AutoJoinTeamCall.ini, Settings, BridgeType
IniRead, adAudioFile  ,%A_ScriptDir%/AutoJoinTeamCall.ini, Settings, AudioFile

VanishingDispMesg("Time for your " . adBridgeType . " meeting, scheduled at " . adConfTime . ". Initiating AutoJoin in " ,2000)
VanishingDispMesg("You can abort AutoJoiner by pressing ESC. Opening link in ",2000)


; Get handle of Teams application
if ( adBridgeType == "Teams" ) {          
   WinGet TEAMSPROG, ID,ahk_exe Teams.exe
   WinActivate ahk_id %TEAMSPROG%
}

 Run %adBridgeUser%    	
 VanishingDispMesg("Waiting to control window and press JOIN button in",	5000)  ; 5 seconds is perhaps too low, but 15 seconds is too long
      
if ( adBridgeType == "Teams" or 0 == 0 ) {            ; Somehow works for Zoom too
  Send {Shift Down}{TAB}{Shift Up}
  Sleep 200
  Send {ENTER}
}

VanishingDispMesg("Waiting for welcome message to end and play Audio file in",6000)
; SoundPlay,  %adAudioFile%   ;;-- Somehow didn't work for big mp3s
Run %adAudioFile%
   
exitapp ;


VanishingDispMesg(text, milliSeconds){
	Gui, +AlwaysOnTop +ToolWindow -SysMenu -Caption
	Gui, Color, ffffff ;changes background color
	Gui, Font, 000000 s18 wbold, Verdana ;changes font color, size and font
	;Gui, Add, Text, x0 y0, %text% 00 Seconds ;the text to display
	;Gui, Show, NoActivate, Xn: 0, Yn: 0

	seconds2Sleep := milliSeconds//1000
	while seconds2Sleep > 0
  {
  	  paddedSeconds2Sleep := (StrLen(seconds2Sleep)=1 ? "0" : "") seconds2Sleep  	  	
    	Gui, Add, Text, x0 y0, %text% %paddedSeconds2Sleep%  Seconds  ;the text to display
	    Gui, Show, NoActivate, Xn: 0, Yn: 0
      seconds2Sleep := seconds2Sleep - 1
      Sleep, 1000
  }
	Gui, Destroy
}


Esc::ExitApp