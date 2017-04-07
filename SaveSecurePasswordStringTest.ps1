
$User = "Mine.Sakiyama@csra.com";
#$SecureStringPassword = Read-Host -AsSecureString "Please enter password"
##$SecureStringPassword = ConvertTo-SecureString "W3dn35d@y321!"  -AsPlainText -Force;
#$SecureStringText = $SecureStringPassword | ConvertFrom-SecureString 
#Set-Content "C:\temp\ExportedPassword.txt" $SecureStringText


$PasswordText = Get-Content "C:\temp\ExportedPassword.txt"
$SecurePassword = $PasswordText| ConvertTo-SecureString 


$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $User, $SecurePassword

# Cleaning up any remaining session. 
Get-PSSession |?{$_.ComputerName -like "*compliance.protection.outlook.com"} | Remove-PSSession

# Connect to the Office 365 Security & Compliance Center https://protection.office.com.
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection 
Import-PSSession $Session 