# Auto Join Teams/Zoom at scheduled time

Tool to schedule joining teams/zoom bridges automatically.

## How ?
There are 3 files in the package

### AutoJoinTeamCallConfig.exe 

Use this to setup the call details, viz. the date-time & the link of the bridge.

![Config Screen](./images/ConfigScreenSample.jpg)

The system datetime is used to populate the date-time fields, while the link & file-path from last execution is reused when the application opens. 

You can also optionally specify a file which will be opened soon after you are connected to the bridge.
You good idea would be to provide path of a mp3 file, recorded with message annoucing your presence.

Once you press the **Schedule Call** button, a task would be schedule using a call to SCHTASKS.EXE.

### AutoJoinCall.exe

This is the supporting application is invoked by **AutoJoinTeamCallConfig** at the scheduled hour.

This application can be terminated by pressing ESC.  It will open the Zoom/Teams link & press the buttons needed to join the call.
Few seconds after the bridge is joined, the file selected usingh AutoJoinTeamCallConfig will be opened.

### AutoJoinTeamCall.ini

This has 2 sections, the "[Settings]" section stores the values set by **AutoJoinTeamCallConfig**. These are used by **AutoJoinTeamCall** when it gets invoked.

The 2nd section "[History]" section has the log of previous SCHTASKS commands. It is a good practice to clean up this section once a while.

## Troubleshooting
