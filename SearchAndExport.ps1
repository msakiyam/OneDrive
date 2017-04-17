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
$PSDFileName = "UserConfig.psd1"
$SecurityAndCompliance = "*compliance.protection.outlook.com"
$SecurityAndComplianceUri = "https://ps.compliance.protection.outlook.com/powershell-liveid/"
$UpdateConfigFile ={  
 
    if(Test-Path $PSDFileName){Remove-Item $PSDFileName}

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
    Set-Content $PSDFileName $PSD1FileString

    Write-Host "`r"
    Write-Host "Done!"
    Write-Host "`r"
    
}

# Check to see if the session exist and usable. 
$ComplianceSession = Get-PSSession |?{$_.ComputerName -like $SecurityAndCompliance}

# Check for Config file to read download dir
if(Test-Path $PSDFileName){
        $Config = Import-LocalizedData -FileName $PSDFileName -BaseDirectory $CurrentWorkingDirectory
        $LocalDirectory = $Config.DownloadDirectory
}
else{
       #If the config file does not exist, run the create config file tool
       Write-Host "The config file was not found. Running the config file tool.." -BackgroundColor White -ForegroundColor Red
       Write-Host "`r"
       &$UpdateConfigFile
       $Config = Import-LocalizedData -FileName $PSDFileName -BaseDirectory $CurrentWorkingDirectory
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
   $ConfigFileExists = Test-Path $PSDFileName
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
   $Config = Import-LocalizedData -FileName $PSDFileName -BaseDirectory $CurrentWorkingDirectory
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

Write-Host "Would you like to proceed to download? Press any key to continue ..." -BackgroundColor Black -ForegroundColor White
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

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


