# Author: Nash King / @gammacapricorni
# By default, this script creates an -AllUserConnection in the public phonebook
# To make a single user connection, find and replace:
#   Remove -AllUserConnection
#   Change $env:PROGRAMDATA to $env:APPDATA
#   Change "$env:Public\Desktop\$ConnectionName.lnk" to [Environment]::GetFolderPath("Desktop") + "\$ConnectionName.link"
#       [Environment]::GetFolderPath("Desktop") allows script to behave correctly on redirected desktops.

# Path for the phonebook.
$PbkPath = Join-Path $env:PROGRAMDATA 'Microsoft\Network\Connections\Pbk\rasphone.Pbk'

# Update these variables with the actual VPN name, address, and PSK.
$ConnectionName = 'VPN name'
$ServerAddress = 'pretend.host.com'
$PresharedKey = 'fake PSK'

# If no VPNs, rasphone.Pbk may not already exist.
# If file does not exist, then create an empty placeholder.
# Placeholder will be overwritten when new VPN is created.
If ((Test-Path $PbkPath) -eq $false) {
    $PbkFolder = Join-Path $env:PROGRAMDATA "Microsoft\Network\Connections\pbk\"
    if ((Test-Path $PbkFolder) -eq $true){
        New-Item -path $PbkFolder -name "rasphone.pbk" -ItemType "file" | Out-Null
    }
    else{
        $ConnectionFolder = Join-Path $env:PROGRAMDATA "Microsoft\Network\Connections\"
        New-Item -path $ConnectionFolder -name "pbk" -ItemType "directory" | Out-Null
        New-Item -path $PbkFolder -name "rasphone.pbk" -ItemType "file" | Out-Null
    }
}

# If VPN exists, delete VPN connection so you can build fresh.
Remove-VpnConnection -AllUserConnection -Name $ConnectionName -Force -EA SilentlyContinue

# Adds the new VPN connection.
Add-VpnConnection -Name $ConnectionName -ServerAddress $ServerAddress -AllUserConnection -TunnelType L2tp -L2tpPsk $PresharedKey -AuthenticationMethod Pap -EncryptionLevel Optional -Force -WA SilentlyContinue

# Sets the VPN connection to split tunnel.
# Comment out for full tunnel.
# Note: Some PCs get angry w/o a short rest to process Add-VPNConnection
Start-Sleep -m 100
Set-VpnConnection -Name $ConnectionName -SplitTunneling $True -AllUserConnection -WA SilentlyContinue

# If you need parameters to add metrics or for IPv6 subnets, open Powershell and run:
# get-help add-vpnconnectionroute -full
# This will give the full list of valid parameters for Add-Vpnconnectionroute and
# instructions for using them.

# Adds the route for the interesting subnet
# $RouteList is an array of interesting subnet(s) with CIDR mask
# Split tunnels must have at least one route.
# Comment out for full tunnel.

$RouteList = @('10.0.1.0/24', '10.1.0.0/16', '10.2.0.0/16')
Foreach ($Destination in $RouteList)
{
    Add-Vpnconnectionroute -Connectionname $ConnectionName -AllUserConnection -DestinationPrefix $Destination
}

# Load the RASphone.pbk file into a line-by-line array
$Phonebook = (Get-Content -path $PbkPath)

# Index for line where the connection starts.
$ConnectionIndex = 0

# Locate the array index for the [$ConnectionName] saved connection.
# Ensures that we only edit settings for this particular connection.
for ($counter=0; $counter -lt $Phonebook.Length; $counter++){
    if($Phonebook[$counter] -eq "[$ConnectionName]"){
        # Set $ConnectionIndex var since $counter only exists inside loop
        $ConnectionIndex = $counter
        break
    }
}

# Starting at the $ConnectionName connection:
# 1. Set connection to use Windows Credential (UseRasCredentials=1)
# 2. Force client to use VPN-provided DNS first (IpInterfaceMetric=1)

# Setting the IpInterfaceMetric to 1 will force the PC to use that DNS first.
# Some companies have local domains that overlap with valid domains
# on the Internet. If VPN-provided DNS can resolve names on the local domain,
# then end user PC will get the correct IP addresses for private servers.
# Otherwise, the PC will use a public DNS resolver.

for($counter=$ConnectionIndex; $counter -lt $Phonebook.Length; $counter++){
    # Set RASPhone.pbk so that the Windows credential is used to
    # authenticate to servers.
    if($Phonebook[$counter] -eq "UseRasCredentials=1"){
        $Phonebook[$counter] = "UseRasCredentials=0"
    }

    # Set RASPhone.pbk so that VPN adapters are highest priority for routing traffic.
    # Comment out if you don't want to try VPN-provided DNS first.
    elseif($Phonebook[$counter] -eq "IpInterfaceMetric=0"){
        $Phonebook[$counter] = "IpInterfaceMetric=1"
        break
        # IpInterfaceMetric comes after UseRasCredentials, so break will cancel
        #   our loop once we're done with it.
    }
}

# Save modified phonebook overtop of RASphone.pbk
Set-Content -Path $PbkPath -Value $Phonebook

# Create desktop shortcut using rasphone.exe.
# Provides a static box for end users to type user name/password into.
# Avoids Windows 10 overlay problems such as showing "Connecting..." even
# after a successful connection.

$ShortcutFile = "$env:Public\Desktop\$ConnectionName.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = "rasphone.exe"
$Shortcut.Arguments = "-d `"$ConnectionName`""
$ShortCut.WorkingDirectory = "$env:SystemRoot\System32\"
$Shortcut.Save()

# Prevent Windows 10 problem with NAT-Traversal (often on hotspots)
# See https://documentation.meraki.com/MX/Client_VPN/Troubleshooting_Client_VPN#Windows_Error_809
# for more details
$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\PolicyAgent"
$name = "AssumeUDPEncapsulationContextOnSendRule"
$value = "2"
New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
