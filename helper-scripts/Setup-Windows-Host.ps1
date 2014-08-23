# Register the HNetCfg library (once)
regsvr32 hnetcfg.dll

#Identify the important adapters
$MainAdapter = Get-NetAdapter | Where-Object {$_.MediaConnectionState -eq 'Connected' -and $_.ComponentID -notlike 'usb*'} | Sort-Object LinkSpeed -Descending | Select-Object -First 1
$EV3Adapter = Get-NetAdapter | Where-Object {$_.MediaConnectionState -eq 'Connected' -and $_.ComponentID -like 'usb*' -and $_.ifDesc -like '*RNDIS*'}

function EnableICS([string]$ID, [int]$Access)
{
    # Create a NetSharingManager object
    $m = New-Object -ComObject HNetCfg.HNetShare

    # List connections
    $m.EnumEveryConnection |% { $m.NetConnectionProps.Invoke($_).Guid }

    # Find connection
    $c = $m.EnumEveryConnection |? { $m.NetConnectionProps.Invoke($_).Guid -eq $ID }
    
    # Get sharing configuration
    $config = $m.INetSharingConfigurationForINetConnection.Invoke($c)
    
    # See if sharing is enabled
    Write-Output $config.SharingEnabled

    # See the role of connection in sharing
    # 0 - public, 1 - private
    # Only meaningful if SharingEnabled is True
    Write-Output $config.SharingType

    # Disable sharing
    #$config.DisableSharing()

    # Enable sharing (0 - public, 1 - private)
    $config.EnableSharing($Access)
}

function DisableAllICS()
{
        # Create a NetSharingManager object
    $m = New-Object -ComObject HNetCfg.HNetShare

    foreach ($c in $m.EnumEveryConnection)
    {    
        # Get sharing configuration
        $config = $m.INetSharingConfigurationForINetConnection.Invoke($c)

        # Disable sharing
        $config.DisableSharing()
    }
}

Write-Output $MainAdapter
Write-Output $EV3Adapter

DisableAllICS
EnableICS $MainAdapter.InterfaceGuid 0
EnableICS $EV3Adapter.InterfaceGuid 1