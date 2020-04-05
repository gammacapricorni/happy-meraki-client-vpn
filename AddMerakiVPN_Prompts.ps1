# Author: Nash King / @gammacapricorni
# By default, this script creates an -AllUserConnection in the public phonebook
# To make a single user connection, do the following:
#   Remove Requires - RunAsAdministrator
#   Remove -AllUserConnection
#   Change $env:PROGRAMDATA to $env:APPDATA
#   Change "$env:Public\Desktop\$ConnectionName.lnk" to [Environment]::GetFolderPath("Desktop") + "\$ConnectionName.link"
#       [Environment]::GetFolderPath("Desktop") allows script to behave correctly on redirected desktops.

#Requires -RunAsAdministrator

# Declare each variable for tidiness.
$ConnectionName = $Continue = $Holder = $PresharedKey = $ServerAddress = $VpnExists = $SplitCheck = $MoreRoutes = ""
$Subnets = @()
$AddRouteCheck = "`nAdd another route? (y/n)"

# Phonebook path for all user connections.
$PbkPath = "$env:PROGRAMDATA\Microsoft\Network\Connections\Pbk\rasphone.pbk"

# If no VPNs, rasphone.pbk may not already exist
# If file does not exist, then create an empty placeholder.

# Placeholder will be overwritten when new VPN is created.
If ((Test-Path $PbkPath) -eq $false) {
    $PbkFolder = "$env:PROGRAMDATA\Microsoft\Network\Connections\pbk\"
    if ((Test-Path $PbkFolder) -eq $true){
        New-Item -path $PbkFolder -name "rasphone.pbk" -ItemType "file" | Out-Null
    }
    else{
        $ConnectionFolder = "$env:PROGRAMDATA\Microsoft\Network\Connections\"
        New-Item -path $ConnectionFolder -name "pbk" -ItemType "directory" | Out-Null
        New-Item -path $PbkFolder -name "rasphone.pbk" -ItemType "file" | Out-Null
    }
}

# Reminder so looping prompts doesn't confuse help desk.
Write-Host -ForegroundColor Yellow "Prompts will loop until you enter a valid response."

# Get VPN connection name.
Do {
    $ConnectionName = Read-Host -Prompt "`nName of VPN Connection"
} While ($ConnectionName -eq "")

# Check if matching VPN already exists.
$VpnExists = (Get-Content $PbkPath | Select-String -Pattern $ConnectionName -Quiet)

# If VPN exists
If ($VpnExists -eq $True) {
    Do {
        # Ask to overwrite
        $Continue = Read-Host -Prompt "`nVPN already exists. Overwrite? (y/n)"
        Switch ($Continue) {
            'y' {
                Try {
                    Remove-VpnConnection -AllUserConnection -Name $ConnectionName -Force
                    Write-Host -ForegroundColor Yellow "`nDeleted old VPN Connection: $ConnectionName"
                }
                Catch {
                    Write-Host -ForegroundColor Red "`nERROR: Unable to delete connection named $ConnectionName"
                    exit
                }
            }
            'n' {
                Write-Host -ForegroundColor Yellow "`nKeeping old VPN. Exiting script..."
                exit
                }
            }
    } Until ($Continue -eq "n" -or $Continue -eq "y")
}

# Use either Meraki dynamic DNS OR a CNAME record of Meraki DDNS for 'host name'.
# Then if IP changes, you don't have to update PCs.
Do {
    $ServerAddress = Read-Host -Prompt "`nHost name or IP address"
} While ($ServerAddress -eq "")

Do {
    $PresharedKey = Read-Host -Prompt "`nPre-shared key"
} While ($PresharedKey -eq "")

# Create the saved VPN connection.
# Suppress error regarding PAP.
Add-VpnConnection -Name $ConnectionName -ServerAddress $ServerAddress -AllUserConnection -TunnelType L2tp -L2tpPsk $PresharedKey -AuthenticationMethod Pap -EncryptionLevel Optional -Force -WA SilentlyContinue
Write-Host -ForegroundColor Yellow "`nCreated VPN connection for $ConnectionName"

# Ask if split or full tunnel
Do {
    $SplitCheck = Read-Host -Prompt "`nSplit tunnel? (y/n)"
    # Prompt for VPN subnet. Ask to confirm. Try to add. Ask to add more routes.
    Switch ($SplitCheck) {
        'y' {
            try {
                Set-VpnConnection -Name $ConnectionName -SplitTunneling $True -AllUserConnection -WA SilentlyContinue
            }
            catch {
                Write-Host -ForegroundColor Red "`nFailed to set split tunnel."
                $SplitCheck = "n"
            }
        }
    }
} Until ($SplitCheck -eq "n" -or $SplitCheck -eq "y")

# If split tunnel, need to add routes for the remote subnets
# Use CIDR format: 192.168.5.0/24
If ($SplitCheck -eq "y") {
    # Loop until at least one valid route is created
    Do {
        # Prompt for the subnet
        Do {
            # Loop until non-blank result given
            Do {
                $Holder = Read-Host -Prompt "`nVPN Subnet"
            } Until ($Holder -ne "")

            # Prompt user to review and approve route
            Do {
                $RouteCheck = Read-Host -Prompt "`nAdd subnet $Holder (y/n)"
            } Until ($RouteCheck -eq "n" -or $RouteCheck -eq "y")

            # If route is approved, try to add
            if ($RouteCheck -eq "y") {
                Try {
                    Add-Vpnconnectionroute -ConnectionName $ConnectionName -AllUserConnection -DestinationPrefix $Holder
                    Write-Host "`nAdded subnet: $Holder"
                    $Subnets += $Holder
                }
                Catch {
                    Write-Host -ForegroundColor Red "`nInvalid route: $Holder."
                    If ($Subnets.count -eq 0) {
                        Write-Host -ForegroundColor Yellow "`nWARNING: No valid subnets have been added to $ConnectionName"
                    }
                }
            }
            $Holder = ""
            # Prompt to add another route
            Do {
                $MoreRoutes = Read-Host -Prompt "$AddRouteCheck"
            } Until ($MoreRoutes -eq "y" -or $MoreRoutes -eq "n")
        # End loop after no more routes
        } While ($MoreRoutes -eq "y")

    # End the loop only once at least one valid subnet has been added
    } Until ($Subnets.count -ge 1)
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
        # Break since we've got our index now
        break
    }
}

# Starting at the $ConnectionName connection:
# 1. Set connection to use Windows Credential (UseRasCredentials)
# 2. Force client to use VPN-provided DNS first (IpInterfaceMetric)
#      Some companies have local domains that overlap with valid domains
#        on the Internet. If VPN-provided DNS can resolve names on the local domain,
#        then end user PC will get the correct IP addresses for private servers.

for($counter=$ConnectionIndex; $counter -lt $Phonebook.Length; $counter++){
    # Set RASPhone.pbk so that the Windows credential is used to
    # authenticate to servers.
    if($Phonebook[$counter] -eq "UseRasCredentials=1"){
        $Phonebook[$counter] = "UseRasCredentials=0"
    }

    # Set RASPhone.pbk so that VPN adapters are highest priority for routing traffic.
    # Comment out if you don't want to use VPN-provided DNS for Internet domains.
    elseif($Phonebook[$counter] -eq "IpInterfaceMetric=0"){
        $Phonebook[$counter] = "IpInterfaceMetric=1"
        break
    }
}

# Save modified phonebook overtop of RASphone.pbk
Set-Content -Path $PbkPath -Value $Phonebook

# Create desktop shortcut for all users using rasphone.exe
# Provides a static box for end users to type user name/password into
# Avoids Windows 10 overlay problems such as showing "Connecting..." even
# after a successful connection.
Try {
    $ShortcutFile = "$env:Public\Desktop\$ConnectionName.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = "rasphone.exe"
    $Shortcut.Arguments = "-d `"$ConnectionName`""
    $ShortCut.WorkingDirectory = "$env:SystemRoot\System32\"
    $Shortcut.Save()
    Write-Host -ForegroundColor Yellow "`nCreated VPN shortcut on desktop for all users. Remind customer to use that short cut!"
}
Catch {
    Write-Host -ForegroundColor Red "`nUnable to create VPN shortcut."
}

# Prevent Windows 10 problem with NAT-Traversal (often on hotspots)
# See https://documentation.meraki.com/MX/Client_VPN/Troubleshooting_Client_VPN#Windows_Error_809
# for more details
$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\PolicyAgent"
$name = "AssumeUDPEncapsulationContextOnSendRule"
$value = "2"
Try {
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
    Write-Host -ForegroundColor Yellow "`nIf this is the first time a Meraki VPN has been setup, reboot computer to finish setup."
}
Catch {
    Write-Host -ForegroundColor Red "`nUnable to create registry key."
}
