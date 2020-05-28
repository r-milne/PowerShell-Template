# PowerShell Template
 PowerShell Template


.SYNOPSIS
Purpose of this script to serve as a template for multiple tasks.  There is logic to loop through multiple Exchange servers.
    
Filters can be adjusted to suit the particular task that is required.  There are multiple examples here:
http://blogs.technet.com/b/rmilne/archive/2014/03/24/powershell-filtering-examples.aspx

An empty array is declared that will be used to hold the data gathered during each iteration. 

This allows for the additional information to be easily added on, and then either echo it to the screen or export to a CSV file

A custom PSObject is used so that we can add data to it from various sources, Get-Mailbox, Get-MailboxStatistics, Get-ADUser etc.
There is no limit to your creativity. 

The CSV is created in the $PWD which is the Present Working Directory, i.e. where the script is saved

 

.DESCRIPTION
This is the template description field.

Replace this description with your own. 

 

.ASSUMPTIONS
 Script is being executed with sufficient permissions to access the server(s) targeted.

 You can live with the Write-Host cmdlets :)

 You can add your error handling if you need it. 

 

.VERSION
  
 1.0  2-12-2014 -- Initial script released to the scripting gallery

 1.1 29--5-2017 - udpated to fix some silly typos

 

 

.Disclaimer
This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. 
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code,
provided that You agree:
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded;
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within the Premier Customer Services Description.
This posting is provided "AS IS" with no warranties, and confers no rights.

Use of included script samples are subject to the terms specified at http://www.microsoft.com/info/cpyright.htm