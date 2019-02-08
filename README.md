# happy-meraki-client-vpn
PowerShell scripts for setting up Meraki Client VPN on Windows 10

Windows 10 doesn't like to play nice with the Meraki client VPN, especially when following Meraki's own setup instructions.

These scripts attempt to:
  1. Pre-emptively fix issues with NAT-Traversal
  2. Simplify creating a split tunnel connection.
  3. Prevent Windows from authenticating to network resources with the VPN credential
  4. Create a rasphone desktop shortcut. Windows 7 users seem to prefer this.
  5. Create the connection for all users. Especially useful for shared laptops or users prone to Windows user profile corruption.
  
  <b>AddMerakiVPN_Prompts.ps1:</b> Handy when you administer multiple Meraki client VPNs, such as at an MSP's help desk. Run in PowerShell and it will prompt you for VPN connection name, VPN concentrator address, pre-shared key, and routes if you pick split tunnel. A new name will create a new connection. If you use a connection name already in use, it will recreate the connection if you indicate. 
  
  <b>AddMerakiVPN.ps1:</b> Edit the variables in the script yourself and then run.
