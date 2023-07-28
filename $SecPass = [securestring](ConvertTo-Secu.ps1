$SecPass          = [securestring](ConvertTo-SecureString "" -AsPlainText -Force)
$LastPassPSCred   = [pscredential](New-Object System.Management.Automation.PSCredential ("josh.lovelace@medhost.com", $SecPass))

Import-Module Lastpass-PS

Connect-Lastpass -Credential $LastPassPSCred

$Acc = Get-Account -Name "LastPass"

Set-Account -Account $Acc -Name $Acc.Name -Credential  $LastPassPSCred -Confirm:$false

ConvertFrom-SecureString -SecureString $Acc.Credential.Password -AsPlainText

New-Password -

Write-Host