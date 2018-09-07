# DLActivityScripts
The following scripts are used to create reports on distribution lists that do not recive emails in your organization.

## Description - DLWeeklyInactiveReport
This script gets all distribution lists in your Office 365 tenant and then runs a message trace on each to see which have been emailed in the past 7 days. It then outputs the emails of the lists that have not recieved email to a text file, this allows the DLMonthlyInactiveReport to compare 4 weeks of results for a monthly report. The script ends by sending an email confirmation that it ran successfully.

## Description - DLMonthlyInactiveReport
This script imports the last 4 weeks of DLWeeklyInactiveReport text file results and compares each to find Distribution Lists that are on each report.  The output is saved in a text file so it can be accessed by the DLQuarterlyInactiveReport. The script then gets details for each of the inactive lists such as display name, primary email, owner, and members. Next, the script checks for weekly report text files older than 5 weeks and removes them - keeping your report folder cleaned up. The details of the inactive lists and removed weekly reports are then formatted in an HTML report that is emailed to you.

## Description - DLQuarterlyInactiveReport
This script imports the last 3 months of DLMonthlyInactiveReport text file results and compares each to find Distribution Lists that are on each report.  The output is saved in a text file so it can be accessed by the DLYearlyInactiveReport. The script then gets details for each of the inactive lists such as display name, primary email, owner, and members. Next, the script checks for monthly report text files older than 4 months and removes them - keeping your report folder cleaned up. The details of the inactive lists and removed monthly reports are then formatted in an HTML report that is emailed to you.

## Requirements
1. Exchange Online PowerShell module is required. Instructions for the module can be found [here] (https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps).

2. ReportHTML module is required. Insctructions for the module can be found [here] (https://www.powershellgallery.com/packages/ReportHTML/).

3. Create a Scheduled task for each of the scripts. The weekly report should run every 7 days. The monthly report should run on the same day as the weekly report every 28 days.  The quarterly report should run on the same day as the daily and monthly reports every 84 days. 
