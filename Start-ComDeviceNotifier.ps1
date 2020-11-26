<#
    .SYNOPSIS
    Runs an event subsriber to notify when a new COM device is available

    .DESCRIPTION
    When the script is running, when you plug (or unplug) a COM device, 
    a notification will pop, telling you a COM device has been plug and what is name is

    .INPUTS
    None.

    .OUTPUTS
    None.

    .EXAMPLE
    C:\PS> .\Update-Month.ps1

    .NOTES
    Uses the module BurntToast to display notification within the notification system of Windows 10

    .NOTES
    The COM Port logo comes from the serialport project website, MIT Licensed
    https://github.com/serialport/website/blob/master/website/static/img/node-serialport-logo-small.svg


#>
[CmdletBinding()]
param (
    # All Device Events
    [Parameter()]
    [switch]
    $AllComDeviceEvent
)
Import-Module BurntToast

$BTHeader = New-BTHeader -Id "ComDeviceNotifier" -Title "Com Device Notifier"

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$NotificationSplat = @{
    Header  = $BTHeader
    AppLogo = $(Join-Path -Path $scriptPath -ChildPath "comlogo.png" )
    Silent  = $true
}
New-BurntToastNotification @NotificationSplat -Text "ComDeviceNotifier Started"

$Action = [scriptblock] {
    param (
        # 
        [Parameter(Position = 0)]
        [object]
        $UselessCimIndicationWatcher,
        # Parameter help description
        [Parameter(Position = 1)]
        [Microsoft.Management.Infrastructure.CimCmdlets.CimIndicationEventInstanceEventArgs]
        $EventInstanceEventArgs
    )
    
    
    switch ($EventInstanceEventArgs.NewEvent.PSTypeNames[0]) {
        "Microsoft.Management.Infrastructure.CimInstance#root/CIMV2/__InstanceCreationEvent" {
            #'Device Arrival' 
            $EventInstanceEventArgs.NewEvent.TargetInstance.Name -match '(?<Name>COM\d)'
            $ComPort = $Matches.Name
            $NotificationSplat.Text = "$ComPort appeared"
        }
        "Microsoft.Management.Infrastructure.CimInstance#ROOT/cimv2/__InstanceDeletionEvent" {
            # 'Device Removal'
            $EventInstanceEventArgs.NewEvent.TargetInstance.Name -match '(?<Name>COM\d)'
            $ComPort = $Matches.Name
            $NotificationSplat.Text = "$ComPort disappeared"
        }
        default { 
            if ($AllComDeviceEvent) {
                $EventInstanceEventArgs.NewEvent.TargetInstance.Name -match '(?<Name>COM\d)'
                $ComPort = $Matches.Name
                $NotificationSplat.Text = "Something happened to $ComPort. You can probably ignore this"
            }
        }
    }

    New-BurntToastNotification @NotificationSplat
}

$query = "Select * FROM __InstanceOperationEvent within 1 where targetInstance isa 'Win32_PnPEntity' and TargetInstance.PNPClass like 'Ports'"
Register-CimIndicationEvent -Query $query -Action $Action -SourceIdentifier "ComDeviceNotifier" | Out-Null

function Stop-ComDeviceNotifier {
    param (    )
    Unregister-Event -SourceIdentifier "ComDeviceNotifier"
    New-BurntToastNotification @NotificationSplat -Text "ComDeviceNotifier Stopped"
}
try {
    while ($true) {
        Wait-Event -Timeout 1
    }
}
catch {
    
}
finally {
    Stop-ComDeviceNotifier
}

