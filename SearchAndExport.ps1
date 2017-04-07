# 
# Title: OneDrive ComplianceSearch and Export
# Author: Mine Sakiyama (Mine.Sakiyama@csra.com)
# Date: 4/4/2017
# Version: 1.2
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
    for($i = 0;$i -le 3;$i++){Write-Host "." -NoNewline; sleep 1}
    Write-Host "`r"

}

# Reading the Config File
$Config = Import-LocalizedData -FileName $PSDFileName -BaseDirectory $CurrentWorkingDirectory

# Retrieving the config information 
$PasswordText = $Config.SecurePasswordString
$SecurePassword = $PasswordText| ConvertTo-SecureString 
$User = $Config.UserName
$LocalDirectory = $Config.DownloadDirectory

# Creating the Cred object
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $User, $SecurePassword

# Cleaning up any remaining session. 
Get-PSSession |?{$_.ComputerName -like "*compliance.protection.outlook.com"} | Remove-PSSession

Write-Host "Connecting to the Seucrity and Compliance Center. Please wait..." -BackgroundColor Black -ForegroundColor White
Write-Host "`r"

Try{
  # Connect to the Office 365 Security & Compliance Center https://protection.office.com.
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection 
  Import-PSSession $Session 
}
Catch{

    $ErrorMessage = $_.Exception.Message
    Write-Host "An error has occured while connecting to the Security and Compliance Center, Please contact your administrator" -ForegroundColor Yellow -BackgroundColor Red
    
}
Finally{
    Write-Host "`r"
    Write-Host "Welcome to the Security and Compliance Center!" -BackgroundColor Black -ForegroundColor White
    
}


# Display the list of Compliace Search to the user and promp for the compliant search name

$Options = @{}
$number = 1
Get-ComplianceSearch |%{
    
    $Options.Add($number,$_.Name)
    $number++

}

$displayOptions = {
   Write-Host "Please select your Compliance Search from the list." 
   for($i=1;$i -le $Options.Count;$i++){

     Write-host $i":" $Options.Item($i)

  }
}

&$displayOptions

$selection = Read-Host "Enter the number"

$CompliantSearchName = $Options.Item([int]$selection)

# Start Compliance Search
""
"Starting the Compliance Search: "  + $CompliantSearchName
Start-ComplianceSearch -Identity $CompliantSearchName

# Check the status
"Please wait for the search to complete..."
Do {

Write-Host "." -NoNewline

}Until((New-ComplianceSearchAction -SearchName $CompliantSearchName -Report -ErrorAction silentlycontinue).Status -eq "Completed")

Write-Host "`r"
"Search Status: " + (New-ComplianceSearchAction -SearchName $CompliantSearchName -Report -ErrorAction silentlycontinue).Status
"Proceeding.."


# Generate the preview 

"Please wait for the preview to complete..."
Do {

Write-Host "." -NoNewline

}Until((New-ComplianceSearchAction -SearchName $CompliantSearchName -Preview -ErrorAction silentlycontinue).Status -eq "Completed")

Write-Host "`r"
"Preview Status: " + (New-ComplianceSearchAction -SearchName $CompliantSearchName -Preview -ErrorAction silentlycontinue).Status
"Retrived the preview, proceeding.."
""

$preview = New-ComplianceSearchAction -SearchName $CompliantSearchName -Preview
$FileCount = 0
[int]$FileSizeTotal = 0
if($preview.Results){
  $preview.Results.Split(";") |%{
  
      if($_ -like "*Size: *"){
  
          $size = [int]($_.split(":")[1]).trim()
          if($size -gt 0){
          $FileCount++
          $FileSizeTotal = $FileSizeTotal + $size
  
          }
  
      }
  
  }

""
"The total number of files to be downloaded: {0}" -f $FileCount
"The total download size {0}MB" -f [math]::Round(($FileSizeTotal/1024/1024),2)
Start-sleep 3
}else{


"Failed to retireve the preview, continuing..."
Start-Sleep 3

}


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


