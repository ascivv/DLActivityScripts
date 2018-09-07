<#	
	.NOTES
	===========================================================================
	 Updated on:   	8/31/2018
	 Created by:    /u/ascIVV
	===========================================================================

        Exchange Online Powershell is required for message trace.
		https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps
		
		ReportHTML Moduile is required
        Install-Module -Name ReportHTML
        https://www.powershellgallery.com/packages/ReportHTML/

	.DESCRIPTION
		Compare three months of DLMonthlyInactivity report results from your O365 tenant. Removes Monthly reports older than 4 months, sends detailed HTML report on unused distribution lists.

#>
#Connection info
$Username = "admin.account@domain.com"
$PasswordPath = "\\path\to\secure\password.txt"

#Read the password from the file and convert to SecureString
$SecurePassword = Get-Content $PasswordPath | ConvertTo-SecureString

#Build a Credential Object from the password file and the $username constant
$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword

#Open a session to O365
$ExOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication  Basic -AllowRedirection
import-PSSession $ExOSession -AllowClobber

#Set the constants
$Date = get-date -format MMddyyyy
$ReportsFolder = "\\server\share\reports\"
$CompanyLogo = "https://www.freelogodesign.org/Content/img/logo-ex-4.png"
$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$RemovedFilesTable = New-Object 'System.Collections.Generic.List[System.Object]'

#This will input past three months of reports and find lists that are on each report. 
#Get report run date for previous weekly reports
$Month1Date = (get-date).AddDays(-56).ToString("MMddyyyy")
$Month2Date = (get-date).AddDays(-28).ToString("MMddyyyy")
$Month3Date = (get-date).ToString("MMddyyyy")

#Clean up monthly reports created more than 4 months ago
$ToOldFiles = ($ReportsFolder+"MonthlyInactive*"+".txt")
$FilestoRemove = Get-ChildItem -Path $ToOldFiles -Force | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-84) }
 
Foreach ($File in $FilestoRemove) {
	Remove-Item $File -Force
		if (!$?) {
		$ReportName =  "$($File.name)"
		$Status 	= "Failed to delete the file automatically."
        }
		
		Else {
		$ReportName = "$($File.name)"
		$Status 	= "Successfully deleted the file automatically."
		}
		
		$obj = [PSCustomObject]@{
		'File Name'		   	   = $ReportName
		'Deleted Status'	   = $Status	
		 }
		
		$RemovedFilesTable.add($obj)
}

If (($RemovedFilesTable).count -eq 0)
{
	$RemovedFilesTable = [PSCustomObject]@{
		'Information'  = 'Information: No Inactive Monthly Lists were found to remove.'
	}
}	

#Set report file path
$Month1Path = ($ReportsFolder+"MonthlyInactive"+$Month1Date+".txt")
$Month2Path = ($ReportsFolder+"MonthlyInactive"+$Month2Date+".txt")
$Month3Path = ($ReportsFolder+"MonthlyInactive"+$Month3Date+".txt")

#Input weekly report files
$Month1Report = Get-Content $Month1Path
$Month2Report = Get-Content $Month2Path
$Month3Report = Get-Content $Month3Path

#Compare weekly report files
$Month12Results =  Compare-Object -ReferenceObject $Month1Report -DifferenceObject $Month2Report -ExcludeDifferent -IncludeEqual
$Month23Results =  Compare-Object -ReferenceObject $Month12Results.InputObject -DifferenceObject $Month3Report -ExcludeDifferent -IncludeEqual

#Filter slider object out of the results
$QuarterlyInactive = $Month23Results.InputObject

#Set export file name for plain text file to be used by yearly report
$DLQuarterlyActivityTxt = ("$ReportsFolder"+"QuarterlyInactive"+"$date"+".txt")

#Export the findings to a file
$QuarterlyInactive | Out-File $DLQuarterlyActivityTxt

#Get inactive distribution list details and create HTML report
Foreach ($List in $QuarterlyInactive) {

		$ListDetails = get-DistributionGroup $List
		$DisplayName = $ListDetails.DisplayName
		$Email = $ListDetails.PrimarySMTPAddress
		$Synced = $ListDetails.IsDirSynced
		$Owner = ($ListDetails.ManagedBy) -join ", "
		$Members = (Get-DistributionGroupMember $List | Sort-Object Name | Select-Object -ExpandProperty Name) -join ", "
		$MeasureMembers = $Members | measure 
		$NumberofMembers = $MeasureMembers.count
		
		$obj = [PSCustomObject]@{
		'Name'				   = $DisplayName
		'Email Address'	       = $Email
		'AD Synced'			   = $Synced
		'Owners'			   = $Owner
		'Members'			   = $Members	
	}
	
	$Table.add($obj)
}

If (($Table).count -eq 0)
{
	$Table = [PSCustomObject]@{
		'Information'  = 'Information: No distribution lists have been inactive for 3 months.'
	}
}

$rpt = New-Object 'System.Collections.Generic.List[System.Object]'
$rpt += get-htmlopenpage -TitleText 'Quarterly Inactive Distribution List Report' -LeftLogoString $CompanyLogo 

		$rpt += Get-HTMLContentOpen -HeaderText "Distribution lists that have not been emailed in 3 months."
            $rpt += get-htmlcontentdatatable $Table -HideFooter
        $rpt += Get-HTMLContentClose
		$rpt += Get-HTMLContentOpen -HeaderText "Monthly Inactive Reports not created in the past 4 months."
		    $rpt += get-htmlcontentdatatable $RemovedFilesTable -HideFooter
	    $rpt += Get-HTMLContentClose
		
$rpt += Get-HTMLClosePage

$rpt += Get-HTMLClosePage
$ReportName = ("DLQuarterlyInactiveReport" + "$Date")
Save-HTMLReport -ReportContent $rpt -ShowReport -ReportName $ReportName -ReportPath $ReportsFolder
$QuarterlyReport = ("$ReportsFolder"+"$ReportName"+".html")

#Send an email with the findings
$From = "admin.account@domain.com"
$To = "helpdesk@domain.com"
$Subject = "Quarterly Distribution List Inactivity Report"
$Body = "See the attached file for distribution lists that have not been emailed in the past 3 months. A .txt file has been saved in the file share to be accessed by the Yearly DL Inactivity Report script. Do not modify any of the weekly, monthly, or quarterly .txt file master copies in the share."
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"

Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Credential -Attachments $QuarterlyReport

#Close the session to O365
Remove-PSSession $ExOSession