#Following class has been downloaded from https://gallery.technet.microsoft.com/scriptcenter/Accessing-Windows-7210ae91
# simple password vault access class
class StoredCredential{ 
  [System.Management.Automation.PSCredential] $PSCredential 
  [string] $account; 
  [string] $password; 

  # loads credential from vault 
  StoredCredential( [string] $name ){ 
      [void][Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]  
      $vault = New-Object Windows.Security.Credentials.PasswordVault  
      $cred = $vault.FindAllByResource($name) | Select-Object -First 1 
      $cred.retrievePassword() 
      $this.account = $cred.userName 
      $this.password = $cred.password  
      $pwd_ss = ConvertTo-SecureString $cred.password -AsPlainText -Force 
      $this.PSCredential = New-Object System.Management.Automation.PSCredential ($this.account, $pwd_ss ) 
  } 

  static [bool] Exists( [string] $name ){ 
      [void][Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]  
      $vault = New-Object Windows.Security.Credentials.PasswordVault  
      try{ 
          $vault.FindAllByResource($name)  
      } 
      catch{ 
          if ( $_.Exception.message -match "element not found" ){ 
              return $false 
          } 
          throw $_.exception 
      } 
      return $true 
  } 

  static [StoredCredential] Store( [string] $name, [string] $login, [string] $pwd ){ 
      [void][Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime] 
      $vault=New-Object Windows.Security.Credentials.PasswordVault 
      $cred=New-Object Windows.Security.Credentials.PasswordCredential($name, $login, $pwd) 
      $vault.Add($cred) 
      return [StoredCredential]::new($name) 
  } 
   
  static [StoredCredential] Store( [string] $name, [PSCredential] $pscred ){ 
      return [StoredCredential]::Store( $name, $pscred.UserName, ($pscred.GetNetworkCredential()).Password ) 
  } 

}

Function Save-O365Credential{
  Param(
    $Mode
  )
  switch($Mode){
    "Basic"{
      $Cred = Get-Credential
      $Cred.Password | ConvertFrom-SecureString | Set-Content "$home\o365-Connection-Password.sec" 
    }
    "Modern"{
      $CredentialName = "Office 365 Connection Script"
      $storedCredential = [StoredCredential]::Store($CredentialName, (Get-Credential))
    }
  }
}

Function ConnectTo-MsolService{

    $username = "admin@foggioncloud.onmicrosoft.com"
    $password = Get-Content "$home\o365-Connection-Password.sec" | convertto-securestring

    If($username -eq $true){
		$global:Credential = Get-Credential
    }
    Else
    {
		$global:Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username,$password
    }

    Connect-MsolService -Credential $global:Credential
}

Function ConnectTo-ExchangeOnlineNew{

    $CredentialName = "Office 365 Connection Script"
    If([StoredCredential]::Exists($CredentialName)){
      $storedCredential = [StoredCredential]::New($CredentialName)
    }
    else{
      $storedCredential = [StoredCredential]::Store($CredentialName, (Get-Credential))
    }

    Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"}|Select-Object -First 1)
    $EXOSession = New-ExoPSSession -Credential $storedCredential.PSCredential
    Import-PSSession $EXOSession
}

Function ConnectTo-ExchangeOnline{

    $username = "admin@foggioncloud.onmicrosoft.com"
    $password = Get-Content "$home\o365-Connection-Password.sec" | convertto-securestring

    If($username -eq $true){
		$global:Credential = Get-Credential
    }
    Else
    {
		$global:Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username,$password
    }

    Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Authentication Basic -Credential $global:Credential -SessionOption (New-PSSessionOption -SkipRevocationCheck -SkipCACheck -SkipCNCheck)  -AllowRedirection)
}

Function ConnectTo-O365SecurityPortal{

    $username = "admin@foggioncloud.onmicrosoft.com"
    $password = Get-Content "$home\o365-Connection-Password.sec" | convertto-securestring

    If($username -eq $true){
		$global:Credential = Get-Credential
    }
    Else
    {
		$global:Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username,$password
	}
	
    Import-PSSession (New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Authentication Basic -Credential $global:Credential -SessionOption (New-PSSessionOption -SkipRevocationCheck -SkipCACheck -SkipCNCheck)  -AllowRedirection) -Prefix "Sec"
}

Function ConnectTo-OfficeCloud{
	ConnectTo-MsolService
	ConnectTo-ExchangeOnlineNew
	ConnectTo-O365SecurityPortal
}