#Schedule this scipts in the Task Scheduler in domain administrator account and trigger on every month first day at 12:00 AM

##################################################################################################################
##                                                                                                              ##
## This scripts used to create the shared mailbox with unlimited storage and apply the fullaccess permission -  ##
## the Exchangegroupadmin usermailbox also trigger the Archive mailbox export job.                            	##
## Created by           : "RAJKUMAR RAMASAMY"                                             			##
## Script creation date : 27th June 2020                                                                        ##
## Last Modified Date   : 27th June 2020                                                                        ##
## Version : 1.0                                                                                                ##
## Make sure the below 11 variable declaration value are correct to execute this job scripts                    ##
##                                                                                                              ##
## Email: rajkumar.@abc.com                                                                 			##
##                                                                                                              ##
##################################################################################################################


#01. $MailboxDB = "Mailbox Database 0798026680"  # Used to store the shared mailbox in which exchange database.
#02. $SMTPMailingdomain = "@domain.com"  # Used to specify the Shared mailbox email account domain.
#03. $ArchiveFilepath1 = "\\172.16.2.51\Archive_PST\"   # Create the shared folder and to specify where the Archived .PST file gets saved.
#04. $logFilePath1 = "C:\Scripts\Log\"    # Create the folder directory and to specify where the script output logs gets saved
#05. $emailSender = "noreply@abc.com"
#06. $emailRecipient = "administrator@abc.com"
#07. $emailCc = "administrator@abc.com"
#08. $emailBcc = "administrator@abc.com"
#09. $emailServer = "smtp.abc.com"
#10. $PSTArchiveJobWaitTime = "1800" #1800 is equal to 30 minutes # Amount of time to complete the archive job
#11. $MailboxToExport = "Journal Account June 2020"  # Specify the Name or identity of the mailbox to export.


$MailboxDB = "MX_DB_JRNL_05"
$logFilePath1 = "C:\Scripts\ShrdMailboxcreation\Log\"
$SMTPMailingdomain = "@domain.com"
$ArchiveFilepath1 = "\\hyper-v10\ExchangeBackup\"
$emailSender = "noreply@abc.com"
$emailRecipient = "Altha@abc.com",Moh@abc.com"
$emailCc = "Kumar@abc.com"
$ArchiveReceipient = "Al@abc.com",Moha@abc.com"
$emailServer = "mail.domain.com"
$PSTArchiveJobWaitTime = "15" #1800 is equal to 30 minutes
$MailboxToExport = "Journal Account June2020"  # Specify the Name or identity of the mailbox to export.

#***********************************************************************************************************************
#***********************************************************************************************************************
Remove-PSSession -ComputerName $hostname
$hostname = [System.Net.Dns]::GetHostByName($env:COMPUTERNAME).HostName
$ExOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$hostname/powershell/ -Authentication Kerberos

Get-Date -UFormat "%A %B/%d/%Y %T %Z"
$CurrentMonth = Get-Date -UFormat "%B"
$CurrentYear = Get-date -UFormat "%Y"
$datetime = get-date -Format yyyyMMdd_hhmmsstt
$date = Get-Date -Format "dd"
Write-Host $date
#Shared Mailbox Parameters like:
$DisplayName = "Journal Account "+$CurrentMonth+" "+$CurrentYear+""
$Alias = "JournalAccount"+$CurrentMonth+""+$CurrentYear+""
$JournalSMTPAddress = "journal"+$CurrentMonth+""+$CurrentYear+"$SMTPMailingdomain"

#$DisplayName = "Journal Account July "+$CurrentYear+""
#$Alias = "JournalAccountJuly"+$CurrentYear+""
#$JournalSMTPAddress = "journaljuly"+$CurrentYear+"$SMTPMailingdomain"

$logFilename = "Log_Email_Archive_$($Alias)_$($dateTime).log"
$logFilePath = $logFilePath1+$logFilename
Write-host $logFilePath


$ArchiveJobName = "Jrnlarchive_"+$Alias+"_"+$datetime+""
Write-Host $ArchiveJobName
$ArchiveFileName = "$Alias.pst"
$ArchiveFilepath = "$ArchiveFilepath1$ArchiveFileName"
write-host $ArchiveFilepath



$EmailSubjectMailboxCreation = "Journal Mailbox Created"
$emailBodyMailboxCreation = "Shared mailbox created $JournalSMTPAddress"

$EmailSubjectmailForwd = "Mailforwarding enabled"
$emailBodymailForwd = "Primary Journal mailbox Emailforwarding enabled to $JournalSMTPAddress"

$emailSubject = "Monthly Email Archive COMPLETE - " + $Alias + " - Log Attached"

$emailBody = "This is an automated message after email archiving scheduled task has completed. Please check attached log file for any problems."

$emailSubjectFailure = "Monthly Email Archive FAILED - " + $Alias + " - Log Attached"

$emailBodyFailure = "This is an automated message after email archiving scheduled task has failed. Please check attached log file for any problems."

$emailAttachment = $logFilePath


#     RECORDING VARILABLES TO LOG FILE

# ----------------------------------------

Write-Output "" | tee $logFilePath -Append
Write-Output "[Variables]" | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$dateTime = ' $dateTime | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveDate = ' $archiveDate | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFile = ' $archiveFile | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFileDir = ' $archiveFileDir | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFilePath = ' $archiveFilePath | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveJobName = ' $archiveJobName | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailSender = ' $emailSender | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailRecipient = ' $emailRecipient | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailCc = ' $emailCc | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailSubject = ' $emailSubject | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailBody = ' $emailBody | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailSubjectFailure = ' $emailSubjectFailure | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailBodyFailure = ' $emailBodyFailure | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailServer = ' $emailServer | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailAttachment = ' $emailAttachment | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$PSScriptRoot = ' $PSScriptRoot | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$logFilePath = ' $logFilePath | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Get-PSSession


Write-Host $hostname

Import-PSSession $ExOPSession
Start-Sleep -Seconds 2

# Create the new shared mailbox with storage limit = unlimited
Write-Output "" | tee $logFilePath -Append
New-Mailbox -Shared -Name $DisplayName -DisplayName $DisplayName -Alias $Alias -Database $MailboxDB -PrimarySmtpAddress $JournalSMTPAddress | Set-Mailbox -UseDatabaseQuotaDefaults $false -IssueWarningQuota unlimited -ProhibitSendQuota unlimited | tee $logFilePath -Append
Write-Output "Mailbox created" | tee $logFilePath -Append
Send-MailMessage -From $emailSender -To $emailRecipient -Cc $emailCc -Subject $EmailSubjectMailboxCreation -Body $emailBodyMailboxCreation -SmtpServer $emailServer -Port 25 -Attachments $emailAttachment

# Update that newly created shared mailbox FullAccess permission to ExchangeGroupAdmin user mailbox
Write-Output "" | tee $logFilePath -Append
Add-MailboxPermission -Identity $DisplayName -User ExchangeGroupadmin -AccessRights FullAccess -InheritanceType All | tee $logFilePath -Append
Write-Output "Mailbox permission updated" | tee $logFilePath -Append
Start-Sleep -Seconds 10

# Update that newly created shared mailbox FullAccess permission to Exchange Journal -IT-Chennai user mailbox
Write-Output "" | tee $logFilePath -Append
Add-MailboxPermission -Identity $DisplayName -User Exchange Journal T-Chennai -AccessRights FullAccess -InheritanceType All | tee $logFilePath -Append
Write-Output "Mailbox permission updated" | tee $logFilePath -Append
Start-Sleep -Seconds 10

#Update the Primary Journal mailbox mailforwarding to that new shared journal mailbox
Write-Output "" | tee $logFilePath -Append
Set-Mailbox -Identity "Exchange.Journal" -DeliverToMailboxAndForward $false -ForwardingAddress $JournalSMTPAddress | tee $logFilePath -Append
Write-Output "Mailforwarding enabled in ' Exchange Journal' $JournalSMTPAddress" | tee $logFilePath -Append
Send-MailMessage -From $emailSender -To $emailRecipient -Cc $emailCc -Subject $EmailSubjectmailForwd -Body $emailBodymailForwd -SmtpServer $emailServer -Port 25 -Attachments $emailAttachment

# Ensure no job of the same name exists.

Get-MailboxExportRequest -Name $archiveJobName | Remove-MailboxExportRequest -confirm:$false | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
#Initiate the .PST mailbox export job 
New-MailboxExportRequest -Mailbox $MailboxToExport -FilePath $ArchiveFilepath -Name $archiveJobName | Set-Notification -NotificationEmails $ArchiveReceipient | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append

# Wait for the archiving job to complete
#Start-Sleep -seconds 10
while (!(Get-MailboxExportRequest -Name $archiveJobName -Status Completed))

{
	Write-Output "" | tee $logFilePath -Append

	Write-Output "Still exporting... Waiting $PSTArchiveJobWaitTime minutes for it to complete..." | tee $logFilePath -Append

	Write-Output "" | tee $logFilePath -Append

	Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl | tee $logFilePath -Append
    #Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl 

	Write-Output "" | tee $logFilePath -Append

	Start-Sleep -seconds $PSTArchiveJobWaitTime
}
$emailArchivingStatus = Get-MailboxExportRequest -Name $archiveJobName -Status Completed | select -expand status

#write-host $emailArchivingStatus
If ($emailArchivingStatus -eq "Completed")
{
    Write-Output "" | tee $logFilePath -Append

	Write-Output "Export completed." | tee $logFilePath -Append

	Write-Output "" | tee $logFilePath -Append

	# Getting more details for manual verification/troubleshooting in log file

	Get-MailboxExportRequest -Name $archiveJobName | fl | tee $logFilePath -Append

	Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl | tee $logFilePath -Append

	# Optional: Clean up (but Get-MailboxExportRequest will no longer return information of the previous job for troubleshooting; instead, check log file for troubleshooting)

	# Get-MailboxExportRequest -Name $archiveJobName | Remove-MailboxExportRequest -confirm:$false
# Send an email using local SMTP server with log when done.

Send-MailMessage -From $emailSender -To $emailRecipient -Cc $emailCc -Subject $emailSubject -Body $emailBody -SmtpServer $emailServer -Port 25 -Attachments $emailAttachment
Remove-PSSession -ComputerName $hostname

}
else

{

	# Report failure via email and exit script

	Write-Output "" | tee $logFilePath -Append

	Write-Output "Export job still inprogress more than 30 minutes" | tee $logFilePath -Append

	Write-Output "" | tee $logFilePath -Append

	# Getting more details for manual verification/troubleshooting in log file

	Get-MailboxExportRequest -Name $archiveJobName | fl | tee $logFilePath -Append

	Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl | tee $logFilePath -Append

	# Optional: Clean up (but Get-MailboxExportRequest will no longer return information of the previous job for troubleshooting; instead, check log file for troubleshooting)

	# Get-MailboxExportRequest -Name $archiveJobName | Remove-MailboxExportRequest -confirm:$false
    $Mailbody1 = "Export job still inprogress more than 30 minutes`n So run this command: Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl`n to knows the export job status "
    Send-MailMessage -From $emailSender -To $emailRecipient -Cc $emailCc -Bcc $emailBcc -Subject "$Mailbody1" -Body $emailBodyFailure -SmtpServer $emailServer -Port 25 -Attachments $emailAttachment
    Remove-PSSession -ComputerName $hostname
	Exit

}


