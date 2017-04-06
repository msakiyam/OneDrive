# 
# Title: OneDrive ComplianceSearch and Export
# Author: Mine Sakiyama (Mine.Sakiyama@csra.com)
# Date: 4/4/2017
# Version: 1.0
# CSRA Inc.(C) All Rights Reserved
# CSRA Think Next Now
# 


#
# User Name and Password below
#
$User = "Mine.Sakiyama@csra.com";
$Password = ConvertTo-SecureString "Tu35d@y54321!"  -AsPlainText -Force;
$Cred = New-Object -TypeName System.Management.Automation.PSCredential($User, $Password);

#
#Enter the name of the Compliance Search Below
#
$CompliantSearchName = "All"

#
#Enter the drive path where you want to downlaod the files to
#
$LocalDirectory = "C:\Temp"

#
#Enter the directory where AZCopy is installed.(Default is already entered below) 
#
$AZCOPYDir = "C:\Program Files (x86)\Microsoft SDKs\Azure\AzCopy"

# Cleaning up any remaining session. 
Get-PSSession |?{$_.ComputerName -like "*compliance.protection.outlook.com"} | Remove-PSSession

# Connect to the Office 365 Security & Compliance Center https://protection.office.com.
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection 
Import-PSSession $Session 


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

#install AZCopy from https://docs.microsoft.com/en-us/azure/storage/storage-use-azcopy
#Path to the AZCopy binaries (Default C:\Program Files (x86)\Microsoft SDKs\Azure\AzCopy needs to be added to the System's Path. 
# /s is required for recursive copy from source dir. 

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
&($CheckSystemPath)
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


