<#	
	.NOTES
	===========================================================================
	 Updated on:   	8/10/2018
	 Created by:   /u/ascIVV
	===========================================================================

        Exchange Online Powershell is required for message trace.
		https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps

	.DESCRIPTION
		Generate a basic report of unused Distrobution Lists in your O365 tenant. Save the results and then compare four weeks of results with the DLMonthlyInactivty report.
    

#>

#Connection info
$Username = "admin.account@domain.com"
$PasswordPath = "\\path\to\secure\password.txt"

#Set the constants
$Date = get-date
$From = "admin.account@domain.com"
$To = @("email.one@domain.com","email.two@domain.com")
$Subject = "Weekly Distribution List Inactivity Report"
$Body = "This is a confirmation the Weekly DL Inactivity Report script has completed successfully. See the attached file for distribution lists that have not been emailed this week. This file has also been saved at " +$DLActivityReport+" so it can be accessed by the Monthly DL Inactivity Report script."
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"

#Set export file name
$DLActivityReport = "\\server\share\DL Activity Reports\Inactive"+$Date.ToString("MMddyyyy")+".txt"

#Read the password from the file and convert to SecureString
$SecurePassword = Get-Content $PasswordPath | ConvertTo-SecureString

#Build a Credential Object from the password file and the $username constant
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword

#Open a session to O365
$ExOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication  Basic -AllowRedirection
import-PSSession $ExOSession -AllowClobber

#This will run a message activity report on all distribution lists for the past week and output the inactive lists.
#Get all Distribution Lists.
$DistroLists = Get-DistributionGroup -ResultSize Unlimited

#Run message trace on each Distribution List to see if it recieved mail in the past x days.
$DistroListsInUse = $DistroLists | select -ExpandProperty primarysmtpaddress  | Foreach-Object { Get-MessageTrace -RecipientAddress $_ -Status expanded -startdate $Date.AddDays(-8) -EndDate $Date -pagesize 1| select -first 1} 

#Check to see if the message trace shows recieved mail vs. not returning anything and output active status respectivly. 
$Results =  Compare-Object -ReferenceObject $DistroListsinUse.RecipientAddress -DifferenceObject $DistroLists.primarySMTPaddress -PassThru

#Export the findings to a file
$Results | Out-File $DLActivityReport

#Send an email with the findings
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Credential -Attachments $DLActivityReport

#Close the session to O365
Remove-PSSession $ExOSession





