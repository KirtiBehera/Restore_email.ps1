param ([string]$Title = "Select opions to restore Emails from mailbox")
Install-Module -Name ExchangeOnlineManagement -Force
Connect-ExchangeOnline 
$email = Read-Host "Enter user Email Address"
[datetime]$date = Read-host "Enter starting Date(mm/dd/yyyy)"
#$resultsize = Read-host "Enter size of reports" (For minimize the ooutput) #Add -resultsize $resultsize for supress the output.
$type = Read-host "Enter (IPM.Note) for Email & (IPM.Appointment) for Meetings & appointments"
$subjects = Read-host "Enter Subject(enter * if no subject available)"
$sender = Read-Host "Enter The Sender's email address(enter * if no sender email available)"

Get-MailboxStatistics $email|Select DisplayName,MailboxType,MailboxTypeDetail,DeletedItemCount,TotalDeletedItemSize

Function Show-Menu { 
Write-Host "============ $Title ============="
Write-Host "1: Display all deleted Emails and Events"
Write-Host "2: Display all deleted Emails and event from Subject"
Write-Host "3: Display all deleted Emails and event from Sender"
Write-Host "4: Restore all selected Emails and Events"
Write-Host "5: Restore all selected Emails and Events from Subject"
Write-Host "6: Restore all selected Emails and from Sender"
Write-Host "Q: Quit."
}

$HtmlHead = '<style>
    body {
        background-color: white;
        font-family:      "Calibri";
    }

    table {
        border-width:     1px;
        border-style:     solid;
        border-color:     black;
        border-collapse:  collapse;
        width:            100%;
    }

    th {
        border-width:     1px;
        padding:          5px;
        border-style:     solid;
        border-color:     black;
        background-color: #98C6F3;
    }

    td {
        border-width:     1px;
        padding:          5px;
        border-style:     solid;
        border-color:     black;
        background-color: White;
    }

    tr {
        text-align:       left;
    }
</style>'

$htmlContent = ""

do {
Show-Menu
$input = Read-Host "Please make a selection"

switch ($input) {
'1' {
cls
Write-Host "You chose option #1" -ForegroundColor Green
Get-RecoverableItems $email -FilterStartTime $date -FilterEndTime $date.adddays(2)  -FilterItemType $type |Select @{Label='USER';Expression={$_.identity}},@{Label='SUBJECT';Expression={$_.Subject}},@{Label='CURRENT FOLDER';Expression={$_.sourcefolder}},@{Label='DELETED FROM';Expression={$_.LastParentPath}},@{Label='ITEM TYPE';Expression={$_.itemclass}},@{Label='LAST UPDATE';Expression={$_.lastmodifiedtime}} 
} 
'2' {
cls
Write-Host "You chose option #2" -ForegroundColor Green
 Get-RecoverableItems $email -SubjectContains $subjects  |Select @{Label='USER';Expression={$_.identity}},@{Label='SUBJECT';Expression={$_.Subject}},@{Label='CURRENT FOLDER';Expression={$_.sourcefolder}},@{Label='DELETED FROM';Expression={$_.LastParentPath}},@{Label='ITEM TYPE';Expression={$_.itemclass}},@{Label='LAST UPDATE';Expression={$_.lastmodifiedtime}}
} 
'3' {
cls
Write-Host "You chose option #3" -ForegroundColor Green
$output=(Get-MessageTrace -SenderAddress $sender -StartDate $date -EndDate $date.adddays(2)).subject 
foreach($subject in $output) {
Get-RecoverableItems $email -SubjectContains $subject|Select @{Label='USER';Expression={$_.identity}},@{Label='SUBJECT';Expression={$_.Subject}},@{Label='CURRENT FOLDER';Expression={$_.sourcefolder}},@{Label='DELETED FROM';Expression={$_.LastParentPath}},@{Label='ITEM TYPE';Expression={$_.itemclass}},@{Label='LAST UPDATE';Expression={$_.lastmodifiedtime}}
}} 

'4' {
cls
Write-Host "You chose option #4" -ForegroundColor Green
$restoreOutput = Restore-RecoverableItems $email -FilterStartTime $date -FilterEndTime $date.adddays(1)  -FilterItemType $type |Select @{Label='USER';Expression={$_.identity}},@{Label='SUBJECT';Expression={$_.Subject}},@{Label='CURRENT FOLDER';Expression={$_.sourcefolder}},@{Label='RESTORED FOLDER';Expression={$_.RestoredToFolderPath}},@{Label='ITEM TYPE';Expression={$_.itemclass}},@{Label='RESTORED STATUS';Expression={$_.WasRestoredSuccessfully}},@{Label='LAST UPDATE';Expression={$_.lastmodifiedtime}}
write-host "We have restored your $type this will take 30 Mins to 60 Mins for replication" -ForegroundColor Green
}
'5' {
cls
Write-Host "You chose option #5" -ForegroundColor Green
$restoreOutput = Restore-RecoverableItems $email -SubjectContains $subjects |Select @{Label='USER';Expression={$_.identity}},@{Label='SUBJECT';Expression={$_.Subject}},@{Label='CURRENT FOLDER';Expression={$_.sourcefolder}},@{Label='RESTORED FOLDER';Expression={$_.RestoredToFolderPath}},@{Label='ITEM TYPE';Expression={$_.itemclass}},@{Label='RESTORED STATUS';Expression={$_.WasRestoredSuccessfully}},@{Label='LAST UPDATE';Expression={$_.lastmodifiedtime}}
Write-host "We have restored your $subjects this will take 30 Mins to 60 Mins for replication" -ForegroundColor Green
} 
'6' {
cls
Write-Host "You chose option #6" -ForegroundColor Green
$output= (Get-MessageTrace -SenderAddress $sender -StartDate $date -EndDate $date.adddays(2)).subject 
foreach ($subject in $output) {
$restoreOutput = Restore-RecoverableItems $email -SubjectContains $subject|Select-Object @{Label='USER';Expression={$_.identity}},@{Label='SUBJECT';Expression={$_.Subject}},@{Label='CURRENT FOLDER';Expression={$_.sourcefolder}},@{Label='RESTORED FOLDER';Expression={$_.RestoredToFolderPath}},@{Label='ITEM TYPE';Expression={$_.itemclass}},@{Label='RESTORED STATUS';Expression={$_.WasRestoredSuccessfully}},@{Label='LAST UPDATE';Expression={$_.lastmodifiedtime}}|ConvertTo-Html -Fragment
$htmlContent += $restoreOutput
}
Write-host "We have restored from Sender:$sender this will take 30 Mins to 60 Mins for replication" -ForegroundColor Green
send-MailMessage -to "kirtiranjan.behera@conteso.com" -from "Restore.Reports@conteso.com" -Cc $email -Subject "Mail Recovery Report $email" -smtpserver "192.168.1.100" -Body "<h1>Mail Recovery Report</h1><br><br>$htmlContent" -BodyAsHtml
Exit
} 
'Q' {
Write-host "You chose option #Q" -ForegroundColor Green 
Write-Host "Thank you for using this Script" -ForegroundColor Yellow
break
}
default {
Write-host "Invalid selection.Please choose again." -ForegroundColor Red
}
}
} until ($input -eq 'Q')

$htmlContent = $restoreOutput|ConvertTo-Html -Head $HtmlHead

Send-MailMessage -to "kirtiranjan.behera@conteso.com" -from "Restore.Reports@conteso.com" -Cc $email -Subject "Mail Recovery Report for $email" -smtpserver "192.168.1.100" -Body "<h1>Mail Recovery Report</h1><br><br>$htmlContent" -BodyAsHtml
