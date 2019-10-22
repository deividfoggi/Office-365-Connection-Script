# Office-365-Connection-Script
This is script can be used in $home\Documents\WindowsPowershell\Profile.ps1 as a easy way to connect to Office 365 Powershell admin modules.

Saving your credentials
 - To save your credentials for basic auth and storing in a .sec file in $home\o365-Connection-Password.sec
      Save-o365Credential -Mode "Basic"
      
 - To save your credentials in credential manager for MFA (Exchange Online Module for MFA). Credential will saved in credential manager as 'Office 365 Connection Script'
      Save-o365Credentials -Mode "Modern"

Connecting to modules

 - Connect to all modules (Exchange Online, Security and Compliance Center, MSOnline Services)
      ConnectTo-OfficeCloud

 - Connec to Exchange Online using basic auth
      ConnectTo-ExchangeOnline

 - Connect to Exchange Online using MFA
      ConnectTo-ExchangeOnlineNew

 - Connect to Security and Compliance portal
      ConnectTo-O365SecurityPortal
