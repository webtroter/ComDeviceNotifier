# COM Device Notifier

PowerShell Script for notifying when a COM device is (un)plugged

## Requirements

This script requires the [BurntToast](https://github.com/Windos/BurntToast) Module ([PowerShell Gallery](http://www.powershellgallery.com/packages/BurntToast/))

The script has been tested on PowerShell 7.1, but should work on 5.1+

## Install

The script must run in the background of the user session. The best option
is to add a scheduled task.

### Scheduled Task

* Add a scheduled Task, name it like you want (ComDeviceNotifier)
* Run only when user is connected
* Trigger : At User Login
  * Since the script is only accessible to my user, I make sure the task only runs when my user is connected
* Action : Start a program
  * Executable : `pwsh.exe`
  * Arguments : `-WindowStyle Hidden -NoProfile -NonInteractive -NoLogo -NoExit -File "C:\Users\vezal\gitrepos\ComDeviceEventNotif\Start-ComDeviceNotifier.ps1"`
* Conditions : I allow the task to be run on battery power
  * Because generally, when I need to know what COM port I just plugged, I'm on the move.
* Parameters : If the task is already running, stop the existing instance.
