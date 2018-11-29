# enumumerate user list with sku info
$filename = "get-userinfo.ps1"
$version = "v1.26 updated on 11/28/2018"
# Jason Himmelstein
# http://www.sharepointlonghorn.com

# Display the profile version
Write-host "$filename $version" -BackgroundColor Black -ForegroundColor Yellow

#region setup session

#PowerShell modules
Install-Module MSOnline
Install-Module -Name AzureAD

#connect to Office 365
Write-host "Connecting to Office 365" -BackgroundColor Yellow -ForegroundColor Black
Connect-MsolService -ErrorAction stop

#Connect to Exchange Online
Write-host "Connecting to Exchange Online" -BackgroundColor Yellow -ForegroundColor Black
Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking 

#endregion

#region output log

$logspath = "C:\RAXLogs"
$FileExists = Test-Path $logspath 
If ($FileExists -eq $True){
write-host "The output location already exist" -ForegroundColor DarkGreen -BackgroundColor Gray}
else{New-Item -Path $logspath -type directory -ErrorAction SilentlyContinue}
        $logname = "{0}\{1}-{2}.{3}" -f $logspath,$env:Username, `
            (Get-Date -Format MMddyyyy-HHmmss),"Txt"
        # Start Transcript in logs directory
        start-transcript (New-Item -Path $logname -ItemType file) -append -noclobber
         $a = Get-Date
        “Date: ” + $a.ToShortDateString()
        “Time: ” + $a.ToShortTimeString() 

#endregion 

#region input file

#specify the input file location
$csvfile = 'users.csv'
$userfile = Read-Host -Prompt "
Enter the location of the CSV file containing the users you want to import. Press Enter for $csvfile"
If ($userfile -eq "") {$userfile = $csvfile}

#endregion

#region execution

#Import the file with the users. You can change the filename to reflect your file
$users = Import-Csv $userfile

#Running for each user
foreach ($user in $users) {
                try {
                    $uUPN=$user.userprinciplename
                    Write-host "Mailbox size information for $uUPN" -BackgroundColor Black -ForegroundColor White
                    Get-Mailbox -Identity $uUPN | Get-MailboxStatistics | Format-Table DisplayName, TotalItemSize, ItemCount -Autosize
                    Write-host "Mailbox information for $uUPN" -BackgroundColor Black -ForegroundColor White
                    get-mailbox -Identity $uUPN -Verbose | fl
                    Write-host "License information for $uUPN" -BackgroundColor Black -ForegroundColor White
                    $AllLicenses=(Get-MsolUser -UserPrincipalName $uUPN).Licenses
                    $licArray = @()
                    for($i = 0; $i -lt $AllLicenses.Count; $i++)
                        {
                        $licArray += "License: " + $AllLicenses[$i].AccountSkuId
                        $licArray +=  $AllLicenses[$i].ServiceStatus
                        $licArray +=  ""
                        }
                    $licArray`
                    }
                catch [System.Object]
                    {
                        Write-Output "Could not get $user info, $_"
                    }
            }

#endregion     
$users | ft
Write-Host "Find the log of these user's information at '$logname'" -BackgroundColor Black -ForegroundColor Green
