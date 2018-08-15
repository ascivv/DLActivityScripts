<#	
	.NOTES
	===========================================================================
	 Updated on:   	8/10/2018
	 Created by:   /u/ascIVV
	===========================================================================

        Exchange Online Powershell is required for message trace.
		https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps

	.DESCRIPTION
		Compare four weeks of DLWeeklyInactivity report results from your O365 tenant.
    

#>


#Connection info
$Username = "admin.account@domain.com"
$PasswordPath = "\\path\to\secure\password.txt"

#Set the constants
$Date = get-date
$From = "admin.account@domain.com"
$To = "helpdesk.email@domain.com"
$Subject = "Monthly Distribution List Inactivity Report"
$Body = "See the attached file for distribution lists that have not been emailed in the past 4 weeks. This file has also been saved at " +$DLMonthlyActivityReport+"  so it can be accessed by the Quarterly DL Inactivity Report script. Please only make modifications to the attached file and not any of the weekly or monthly master copies in the share."
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"

#Read the password from the file and convert to SecureString
$SecurePassword = Get-Content $PasswordPath | ConvertTo-SecureString

#Build a Credential Object from the password file and the $username constant
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword

#Open a session to O365
$ExOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication  Basic -AllowRedirection
import-PSSession $ExOSession -AllowClobber

#This will input past four weeks of reports and find lists that are on each report. 
#Get report run date for previous weekly reports
$Week1Date = $Date.AddDays(-21).ToString("MMddyyyy")
$Week2Date = $Date.AddDays(-14).ToString("MMddyyyy")
$Week3Date = $Date.AddDays(-7).ToString("MMddyyyy")
$Week4Date = $Date.ToString("MMddyyyy")

#Set report file path
$Week1Path = "\\server\share\DL Activity Reports\Inactive\Inactive"+$Week1Date+".txt"
$Week2Path = "\\server\share\DL Activity Reports\Inactive\Inactive"+$Week2Date+".txt"
$Week3Path = "\\server\share\DL Activity Reports\Inactive\Inactive"+$Week3Date+".txt"
$Week4Path = "\\server\share\DL Activity Reports\Inactive\Inactive"+$Week4Date+".txt"

#Input weekly report files
$Week1Report = Get-Content $Week1Path
$Week2Report = Get-Content $Week2Path
$Week3Report = Get-Content $Week3Path
$Week4Report = Get-Content $Week4Path

#Compare weekly report files
$Week12Results =  Compare-Object -ReferenceObject $Week1Report -DifferenceObject $Week2Report -ExcludeDifferent -IncludeEqual
$Week23Results =  Compare-Object -ReferenceObject $Week12Results.InputObject -DifferenceObject $Week3Report -ExcludeDifferent -IncludeEqual
$Week34Results =  Compare-Object -ReferenceObject $Week23Results.InputObject -DifferenceObject $Week4Report -ExcludeDifferent -IncludeEqual

#Filter slider object out of the results
$MonthlyInactive = $Week34Results.InputObject

#Set export file name
$DLMonthlyActivityReport = "\\server\share\DL Activity Reports\Inactive\MonthlyInactive"+$Date.ToString("MMddyyyy")+".txt"

#Export the findings to a file
$MonthlyInactive | Out-File $DLMonthlyActivityReport

#Send an email with the findings
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Credential -Attachments $DLMonthlyActivityReport

#Close the session to O365
Remove-PSSession $ExOSession
