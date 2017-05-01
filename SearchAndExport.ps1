# 
# Title: OneDrive ComplianceSearch and Export
# Author: Mine Sakiyama (Mine.Sakiyama@csra.com)
# Date: 4/17/2017
# Version: 2.0
# CSRA Inc.(C) All Rights Reserved
# CSRA Think Next Now
# 
#
# This script requires that you have installed a tool called AZCopy.
#
# Install AZCopy from https://docs.microsoft.com/en-us/azure/storage/storage-use-azcopy
#
# Enter the directory where AZCopy is installed.(Default is already entered below) 
#
$AZCOPYDir = "C:\Program Files (x86)\Microsoft SDKs\Azure\AzCopy"

$CurrentWorkingDirectory = (Get-Item -Path ".\" -Verbose).FullName
$PSDFileDir = "$env:HOMEDRIVE$env:HOMEPATH"
$PSDFileName = "CSRAUserConfig.psd1"
$PSDFileFullPath = "$PSDFileDir\$PSDFileName"
$SecurityAndCompliance = "*compliance.protection.outlook.com"
$SecurityAndComplianceUri = "https://ps.compliance.protection.outlook.com/powershell-liveid/"

$UpdateConfigFile ={  
 
    if(Test-Path $PSDFileFullPath){Remove-Item $PSDFileFullPath}

    Write-Host "`r"
    Write-Host "Welcome to OneDrive Scritp Configuration Tool.." -BackgroundColor Black -ForegroundColor White
    Write-Host "Press any key to continue ..." -BackgroundColor Black -ForegroundColor White
    Write-Host "`r"

    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    $UserName = Read-Host "Please enter your CSRA email address."
    Write-Host "`r"
    $SecureStringPassword = Read-Host -AsSecureString "Please enter password"
    Write-Host "`r"
    
    $DownloadDir = Read-Host "Please enter the directory where you want to download the exported items"
    Write-Host "`r"

    while(!(Test-Path $DownloadDir)){

     $DownloadDir = Read-Host "The directory you specified was not found.. Please enter the directory where you want to download the exported items"

     Write-Host "`r"

    }

    $SecureStringText = $SecureStringPassword | ConvertFrom-SecureString 
    
    #PSD1 Format 
    #@{UserName = 'mine.sakiyama@csra.com';SecurePasswordString = ""}
    
    $PSD1FileString = '@{UserName = ' + "`'" + $UserName + "`'" + ";"
    $PSD1FileString = $PSD1FileString + "DownloadDirectory = " + "`"" + $DownloadDir + "`";"
    $PSD1FileString = $PSD1FileString + "SecurePasswordString = " + "`"" + $SecureStringText + "`"}"
    
    Write-Host "Generating the config file, please wait..."
    for($i = 0;$i -le 2;$i++){Write-Host "." -NoNewline; sleep 1}

    # Updating the Config File
    Set-Content $PSDFileFullPath $PSD1FileString

    Write-Host "`r"
    Write-Host "Done!"
    Write-Host "`r"
    
}

# Check to see if the session exist and usable. 
$ComplianceSession = Get-PSSession |?{$_.ComputerName -like $SecurityAndCompliance}

# Check for Config file to read download dir
if(Test-Path $PSDFileFullPath){
        $Config = Import-LocalizedData -FileName $PSDFileName -BaseDirectory $PSDFileDir
        $LocalDirectory = $Config.DownloadDirectory
}
else{
       #If the config file does not exist, run the create config file tool
       Write-Host "The config file was not found. Running the config file tool.." -BackgroundColor White -ForegroundColor Red
       Write-Host "`r"
       &$UpdateConfigFile
       $Config = Import-LocalizedData -FileName $PSDFileName -BaseDirectory $PSDFileDir
       $LocalDirectory = $Config.DownloadDirectory
}

#Check for live session
if($ComplianceSession.State -eq "Opened" -and $ComplianceSession.Availability -eq "Available"){
# If the session is still good, do nothing...
}
else{
   
   Write-Host "No live session found.. Creating new session..." -BackgroundColor White -ForegroundColor Red
   #remove any remaining session
   Get-PSSession |?{$_.ComputerName -like $SecurityAndCompliance} | Remove-PSSession -ErrorAction SilentlyContinue | Out-Null
     
   #Checking for the config file
   $ConfigFileExists = Test-Path $PSDFileFullPath
   if(!$ConfigFileExists){
   
       #If the config file does not exist, run the create config file tool
       Write-Host "The config file was not found. Running the config file tool.." -BackgroundColor White -ForegroundColor Red
       Write-Host "`r"
       &$UpdateConfigFile
   
   }else{
       
       Write-Host "`r"
       Write-Host "The config file found. Reading the config file ..." -BackgroundColor Black -ForegroundColor White
       for($i = 0;$i -le 3;$i++){Write-Host ".." -NoNewline; sleep 0.1}
       Write-Host "`r"
   
   }
   
   #genereate cred object for login.
   #region Reading the Credential File
   $Config = Import-LocalizedData -FileName $PSDFileName -BaseDirectory $PSDFileDir
   $LocalDirectory = $Config.DownloadDirectory
   # Retrieving the config information 
   $PasswordText = $Config.SecurePasswordString
   $SecurePassword = $PasswordText| ConvertTo-SecureString 
   $User = $Config.UserName
   $LocalDirectory = $Config.DownloadDirectory
   # Creating the PS Cred object
   $Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $User, $SecurePassword
   
   # Creating SharePoint Online Credential
   #$SecureString = ConvertTo-SecureString -String $Password -AsPlainText -Force
   #$SPOCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$SecurePassword)
   
   #endregion
   
    Try{
          # Connect to the Office 365 Security & Compliance Center https://protection.office.com.
          $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $SecurityAndComplianceUri -Credential $Cred -Authentication Basic -AllowRedirection 
          Import-PSSession $Session 
        }
    Catch{
        
            $ErrorMessage = $_.Exception.Message
            Write-Host "An error has occured while connecting to the Security and Compliance Center, Please contact your administrator" -ForegroundColor Yellow -BackgroundColor Red
            
        }
    Finally{

            Write-Host "`r"
            Write-Host "Welcome to the Security and Compliance Center!" -BackgroundColor Black -ForegroundColor White
            Write-Host "`r"
        }
        
}

# Display the list of eDiscovery Case to the user and promp for the eDiscovery Case Name

$Options = @{}
$number = 1
Get-ComplianceCase |%{
 $CaseName =  $_.Name
 $CaseID = $_.Identity
 $CaseNameandID = @()
 $CaseNameandID += $CaseName
 $CaseNameandID += $CaseID
 $Options.Add($number,$CaseNameandID)
 $number++

 #$CaseName = $_.Name
 #$CaseID = $_.Identity
 #$Search = Get-ComplianceSearch -Case $CaseID
 ##$CaseName + "`t"  + $Search.Name + "`t" + $Search.RunBy + "`t" + $Search.JobEndTime + "`t" + $Search.Status + "`t" + $Search.Size
 #Write-Host $CaseName " " -BackgroundColor White -ForegroundColor Black -NoNewline
 #Write-host "(" $Search.Name ")" -BackgroundColor Gray -ForegroundColor Blue

}

Function CheckOption {
 Param($userselection,
       [int]$numberOfOptions)
   if(($userselection -as [int]) -is [int] -and ($userselection -as [int]) -le $numberOfOptions -and ($userselection -as [int]) -ne 0){
   return $true
   }else{
   return $false
   }
}
$displayOptions = {
     Write-Host "`r"
     Write-Host "Please select your eDiscovery case from the list." 
     ""
     for($i=1;$i -le $Options.Count;$i++){
  
       Write-host $i":" $Options.Item($i)[0]
  
    }
  }
Do{
  
  &$displayOptions
  ""
  $selection = Read-Host "Enter the number"

}while((CheckOption -userselection $selection -numberOfOptions $Options.Count) -eq $false)

$eDiscoveryCaseName = $Options.Item([int]$selection)[0]
$eDiscoveryCaseID = $Options.Item([int]$selection)[1]

#Look for a search assoicated with the case ID
$Search = Get-ComplianceSearch -Case $eDiscoveryCaseID -ErrorAction silentlycontinue

if($Search){
   # Start Compliance Search
   ""
   "The following Search was found:"
   $CompliantSearchName = $Search.Name
   Write-host $CompliantSearchName -BackgroundColor white -ForegroundColor black
  
    #$Search.Name + "`t" + $Search.RunBy + "`t" + $Search.JobEndTime + "`t" + $Search.Status + "`t" + $Search.Size
   Write-host "Starting the Compliance Search: " -NoNewline
   Write-Host  $CompliantSearchName -BackgroundColor white -ForegroundColor Black
   Write-Host "`r"
   Write-Host "Press any key to continue ..." -BackgroundColor Black -ForegroundColor White
   Write-Host "`r"

    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
   
   #kick off the search...
   Start-ComplianceSearch -Identity $Search.Identity

}else{

  Write-Host "No associated search found!"

}


# Check the status of the search...
Write-Host "Running the Search..."
$hourglassCounter = 0
$PleaseWait = "Please wait:"
$HourGlass = @("`|","`/","`-","`\")
$y = $CursorTop=[Console]::CursorTop
$x = $PleaseWait.Length
Write-Host $PleaseWait -NoNewline -ForegroundColor Yellow -BackgroundColor Blue
Do { 
    [Console]::SetCursorPosition($x,$y)
    Write-Host $HourGlass[$hourglassCounter] -ForegroundColor Yellow -BackgroundColor Blue -NoNewline
    sleep 0.2
    if($hourglassCounter -lt $HourGlass.Count){ $hourglassCounter++ }else{$hourglassCounter = 0}
       
}until((Get-ComplianceSearch -Identity (Get-ComplianceSearch -Case $eDiscoveryCaseID -ErrorAction silentlycontinue).Identity).Status -eq "Completed")

Write-Host "`r"

#report the number of files and total size of the download
$Search = Get-ComplianceSearch -Identity (Get-ComplianceSearch -Case $eDiscoveryCaseID -ErrorAction silentlycontinue).Identity

Write-Host "Total Number of Items:  " -BackgroundColor White -ForegroundColor Black -NoNewline
Write-host $Search.Items -ForegroundColor Black -BackgroundColor Gray
Write-Host "Total Size: " -BackgroundColor White -ForegroundColor Black -NoNewline
write-host ([math]::Round(($Search.Size/1024/1024),2)) "MB" -ForegroundColor Black -BackgroundColor Gray

if([int]$Search.Items -ge 1){

    Write-Host "Would you like to proceed to download? Press any key to continue ..." -BackgroundColor Black -ForegroundColor White
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

} else {

    Write-Host "The search did not find any items to download.. exiting" -BackgroundColor Green -ForegroundColor Yellow
        
    Break

}


#region Report and Preview (propbably not needed)
## Check the status
#"Please wait for the search to complete..."
#Do {
#
#Write-Host "." -NoNewline
#
#}Until((New-ComplianceSearchAction -SearchName $CompliantSearchName -Report -ErrorAction silentlycontinue).Status -eq "Completed")
#
#Write-Host "`r"
#"Search Status: " + (New-ComplianceSearchAction -SearchName $CompliantSearchName -Report -ErrorAction silentlycontinue).Status
#"Proceeding.."
#
#
## Generate the preview 
#
#"Please wait for the preview to complete..."
#Do {
#
#Write-Host "." -NoNewline
#
#}Until((New-ComplianceSearchAction -SearchName $CompliantSearchName -Preview -ErrorAction silentlycontinue).Status -eq "Completed")
#
#Write-Host "`r"
#"Preview Status: " + (New-ComplianceSearchAction -SearchName $CompliantSearchName -Preview -ErrorAction silentlycontinue).Status
#"Retrived the preview, proceeding.."
#""
#
#$preview = New-ComplianceSearchAction -SearchName $CompliantSearchName -Preview
#$FileCount = 0
#[int]$FileSizeTotal = 0
#if($preview.Results){
#  $preview.Results.Split(";") |%{
#  
#      if($_ -like "*Size: *"){
#  
#          $size = [int]($_.split(":")[1]).trim()
#          if($size -gt 0){
#          $FileCount++
#          $FileSizeTotal = $FileSizeTotal + $size
#  
#          }
#  
#      }
#  
#  }
#
#""
#"The total number of files to be downloaded: {0}" -f $FileCount
#"The total download size {0}MB" -f [math]::Round(($FileSizeTotal/1024/1024),2)
#
#
#Start-sleep 2
#
#}else{
#
#"Failed to retireve the preview, continuing..."
#Start-Sleep 3
#
#}
#endregion

# Export the result to Azure Drive

$export = New-ComplianceSearchAction $CompliantSearchName -Export -IncludeCredential

# Retireve the Blob URL and SAS token from the Export Result. 

$BlobUrl = "https:" + ((($export.Results).Split(";")[0]).split(":")[2]).trim()
$SASToken = ((($export.Results).Split(";")[1]).split(":")[1]).trim()

# Prep for AZCopy
$CheckSystemPath = {
   $count = 0
   $env:Path.Split(";") |%{
   
       if($_.toString() -eq $AZCOPYDir){
       $count++
       }
   
   }
      if($count -gt 0){
           return $true
      }else{
           return $false
      }
}

if(!(&($CheckSystemPath))){
    $AZCopyPATH = ";$AZCOPYDir"
    $env:Path += $AZCopyPATH
}

$JournalFile = "$Home\AppData\Local\Microsoft\Azure\AzCopy\AzCopy.jnl"

If(Test-Path $JournalFile){Remove-Item $JournalFile}
""
"Starting download..."
""

AZCOPY /Source:$BlobUrl$SASToken /s /v /CheckMD5 /Dest:$LocalDirectory

# SIG # Begin signature block
# MIIMlwYJKoZIhvcNAQcCoIIMiDCCDIQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUGDuhTllGpjIi2BnbO788T1Ru
# tv6gggnjMIIEQTCCAymgAwIBAgITdAAAAAI2HNGawWreLAAAAAAAAjANBgkqhkiG
# 9w0BAQsFADAiMSAwHgYDVQQDExdDU1JBIEVudGVycHJpc2UgUm9vdCBDQTAeFw0x
# NTEyMTYxODU4MzNaFw0yMTEyMTYxOTA4MzNaMGYxEzARBgoJkiaJk/IsZAEZFgNj
# b20xFDASBgoJkiaJk/IsZAEZFgRjc3JhMRQwEgYKCZImiZPyLGQBGRYEY29ycDEj
# MCEGA1UEAxMaQ1NSQSBFbnRlcnByaXNlIElzc3VpbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQDVIQCJO/HOC8mtyAXGTwSsAa9nQWSRtblLyL4P
# O9h760RkYpsagkvyMNrSu6kPb0iuhdaprcwDXUUSYqZpn3MVViB19AZE8mF37z4n
# eakT6DTAddcR5W0hLM9zUzREgUCRF6DeW/8oylSKrOMRT5IDOV7tO/m8hx1AZs0K
# eZYnz+GPH+f6FNLPXkLEStcdg0OGklhUxnPthGOwj5H98On6NcFh/U4iPWvMo+9X
# C7hOxfXVkjZcQ1UX5wJc5XTtnV1lJsphTcxrdFpoduzbOGo+fDuxkG1GyzuHj8X/
# h8m0E1TVHMtoQztgBTuy4BGxGOGT1fLzjTB75LFQTuObukplAgMBAAGjggEqMIIB
# JjAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQULcarIkxTJE+Ld/CJ0awzRxaY
# /dYwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAUQMIFsspNG1dqqZrEBcmLM6HDdVgwRgYD
# VR0fBD8wPTA7oDmgN4Y1aHR0cDovL3BraS5jc3JhLmNvbS9DZXJ0ZGF0YS9DU1JB
# RW50ZXJwcmlzZVJvb3RDQS5jcmwwUQYIKwYBBQUHAQEERTBDMEEGCCsGAQUFBzAC
# hjVodHRwOi8vcGtpLmNzcmEuY29tL0NlcnREYXRhL0NTUkFFbnRlcnByaXNlUm9v
# dENBLmNydDANBgkqhkiG9w0BAQsFAAOCAQEAEpEd3ssqDAi56mEagAil2wqK8ZFo
# htoDceDsQ+R2SmKDFfxMqcyMmlg+rDoxVlH7TOrprSvbz399IRgHHYBNouFoloVT
# DOU8s776AkVepsVn/IiVRObM4+FpKxJ7t/PE2dAmf04hqkb/3IorZwHAc49SDB4r
# 1acIL6mtJZ71cIjlVsI23a0gMdVSAySDYjXJs1W0tBnSDZA7fHNZimix3OWUIXIs
# r3a8becXtPFqtqAtS2SFRBko+7l2vDq+SZtgEnNt9873O5gd2h1FXv5KeP1ee1kU
# gClguO1EmulcceS1euryu5V6TBidVZSefljP15zZRLiezQbxNzvKhgfs9TCCBZow
# ggSCoAMCAQICExIAACDLOw/VdYqH2UcAAAAAIMswDQYJKoZIhvcNAQELBQAwZjET
# MBEGCgmSJomT8ixkARkWA2NvbTEUMBIGCgmSJomT8ixkARkWBGNzcmExFDASBgoJ
# kiaJk/IsZAEZFgRjb3JwMSMwIQYDVQQDExpDU1JBIEVudGVycHJpc2UgSXNzdWlu
# ZyBDQTAeFw0xNzA0MTcyMTQzMzlaFw0yMDA0MTYyMTQzMzlaMIGrMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRQwEgYKCZImiZPyLGQBGRYEY3NyYTEUMBIGCgmSJomT8ixk
# ARkWBGNvcnAxFTATBgNVBAsTDFVzZXJzLUdyb3VwczERMA8GA1UECxMIQXBwLVBy
# aXYxFzAVBgNVBAMTDlNha2l5YW1hLCBNaW5lMSUwIwYJKoZIhvcNAQkBFhZNaW5l
# LlNha2l5YW1hQGNzcmEuY29tMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAz+U9SadzUvXAa9v+fHBJGcfG4kw6wxz08wAhDF+cWTJjChuTQ4N/nXnMhAVk
# 3iEykVbRn5JiO27t4FxWlI2hAwrQJbus8WhG0v5G7/kEObzOPlBjmq7dc8Zagt9J
# tkXfGHeiC5o5WgwquwW1fa5Yl74IVUf8ARv5YYXnTzFK4XtjjCdyy9eW1fhLmk5Y
# cpq8PGWxctynKslg+257aSCDO0cJwnw8PiXkcyk6EpD07li99miO6Si1pMX6QbRQ
# NQuIKV4/vgAwn61Ip79Ip+uOE3Qs/KONbBUbe24nNI83k2aK05NcIBXoAyzWhQjR
# Sqbe+5y9GYH+W0d+EfSzwko97QIDAQABo4IB+TCCAfUwPQYJKwYBBAGCNxUHBDAw
# LgYmKwYBBAGCNxUIhI36EoHEvyWDhYMegrL8N4PEl25NhPy5XIeHoCoCAWQCAQww
# EwYDVR0lBAwwCgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgWgMBsGCSsGAQQBgjcV
# CgQOMAwwCgYIKwYBBQUHAwMwRAYJKoZIhvcNAQkPBDcwNTAOBggqhkiG9w0DAgIC
# AIAwDgYIKoZIhvcNAwQCAgCAMAcGBSsOAwIHMAoGCCqGSIb3DQMHMB0GA1UdDgQW
# BBQqCCnMYkjTvu2v7CxhcMAOsfxQXTAfBgNVHSMEGDAWgBQtxqsiTFMkT4t38InR
# rDNHFpj91jBJBgNVHR8EQjBAMD6gPKA6hjhodHRwOi8vcGtpLmNzcmEuY29tL2Nl
# cnRkYXRhL0NTUkFFbnRlcnByaXNlSXNzdWluZ0NBLmNybDBUBggrBgEFBQcBAQRI
# MEYwRAYIKwYBBQUHMAKGOGh0dHA6Ly9wa2kuY3NyYS5jb20vY2VydGRhdGEvQ1NS
# QUVudGVycHJpc2VJc3N1aW5nQ0EuY3J0MEsGA1UdEQREMEKgKAYKKwYBBAGCNxQC
# A6AaDBh4LW1zYWtpeWFtQGNvcnAuY3NyYS5jb22BFk1pbmUuU2FraXlhbWFAY3Ny
# YS5jb20wDQYJKoZIhvcNAQELBQADggEBAEhLtnOZ/wxS65YbEDXt8O02enTQ9VBG
# KP2Kz4S/wicpLHorJ7l0Z5Kv6fR4qF8WA0NIQUWRcYKbPZs6P6kR1j+qBD/wbds3
# AtdkCuiy2cunRsXszCtkuCkKmHYxpNKXbztNv3dRlXe3Hv+BBL3SkdEqpu+1dM+v
# ZLbXsQvmEwncHCqdjNklAMdfcW8uuAccoZlYDd8t3ckOnXxf/pO/2Eo2/U6HynMn
# jAUyH/6VJ9z9ba+dQTDbQJHH1NWrOtG6qLW2+jJAlouxivCrggpUjTYzejrK9tYo
# E79cJFsu9bKF87vDws3SUHNQiA4127xI0P/8eKk92Wx2WVtUKFC1zU0xggIeMIIC
# GgIBATB9MGYxEzARBgoJkiaJk/IsZAEZFgNjb20xFDASBgoJkiaJk/IsZAEZFgRj
# c3JhMRQwEgYKCZImiZPyLGQBGRYEY29ycDEjMCEGA1UEAxMaQ1NSQSBFbnRlcnBy
# aXNlIElzc3VpbmcgQ0ECExIAACDLOw/VdYqH2UcAAAAAIMswCQYFKw4DAhoFAKB4
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFB9harkRy0a+oXH3Z8cxNqLGm7dIMA0GCSqGSIb3DQEBAQUABIIBAIuVNNWQ
# VfQtRtZFG3dzVvuwA1K5WRfTNzFc9j9jxvXSUVzvp3k8tpjzk0Uxgtb/gQhl/BVK
# 2Ap8X1Lsz3T2IMgXeWKg1Swh0pjCN94d5yaWpSmTRnHodO3Wb/2UEGNLT2Vcccyd
# eKE0B0U2ThMxmlidmHvbfgsVR9gPKBjcZN0arDJZvjRD3cJ8dE/phC7BITCproR7
# /wV2Ckxa15ziASr9Z/Zba5cjHMrGVF+9we9e+JT6BYipgvsAHn+xIkuSKoMr6wMN
# ddawcoWMKpwKHhoCG5bOxSWxgzRCtGshDJIDLwcLe2Lrlww4oXrgTWz35elywChe
# tiwd4/PdI5qKavE=
# SIG # End signature block
