<# 

.SYNOPSIS
	Purpose of this script to serve as a template for multiple tasks.  There is logic to loop through multiple Exchange servers.  That can be ammended as necessary 
    
    Filters can be adjusted to suit the particular task that is required.  There are multiple examples here:
    https://blog.rmilne.ca/2014/03/24/exchange-powershell-filtering-examples


    An empty array is declared that will be used to hold the data gathered during each iteration. 
    This allows for the additional information to be easily added on, and then either echo it to the screen or export to a CSV file 

    A custom PSObject is used so that we can add data to it from various sources, Get-Mailbox, Get-MailboxStatistics, Get-ADUser etc.
    There is no limit to your creativity!  

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
	1.1  29-05-2017 -- Updated script to fix some silly typos 
	1.2  28-5-2020  --  Updated to new blog URL 

    



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

Use of included script samples are subject to the terms specified at http://www.microsoft.com/info/cpyright.htm.

#>



# Declare an empty array to hold the output
$Output = @()

# Declare a custom PS object. This is the template that will be copied multiple times. 
# This is used to allow easy manipulation of data from potentially different sources 
# Element1, Element2 etc are all arbitary names.  You can call them what you want.  Just use the name consistently.... 
# For example we could call each entery Element1 and Element2 and Element3
# $TemplateObject = New-Object PSObject | Select Element1, Element2, Element3
#
# Realistically the below is a real example of what we could use this for.  Note that these names are aligned with the usage in the loop below.  
$TemplateObject = New-Object PSObject | Select ServerName, ServerVersion, ServerInstallDate

# Get a list of Exchange 2010 CAS servers.  Sort this based on name to meet my OCD requirements...
# Can easily edit the query.
# 
# For more filtering examples see: 
# https://blog.rmilne.ca/2014/03/24/exchange-powershell-filtering-examples/ 
#  
$ExchangeServers = Get-ExchangeServer | Where-Object {$_.AdminDisplayVersion -match "^Version 14" -and $_.ServerRole -Match "ClientAccess" }  | Sort-Object  Name




# Loop me baby! 
ForEach ($Server In $ExchangeServers)
{
	# Make a copy of the TemplateObject.  Then work with the copy...
    $WorkingObject = $TemplateObject | Select-Object * 

    # Populate the TemplateObject with the necessary details.
    $WorkingObject.ServerName    = $Server.Name 
    $WorkingObject.ServerVersion = $server.AdminDisplayVersion 
    $WorkingObject.ServerInstallDate = $Server.WhenCreated


    # Display output to screen.  REM out if not reqired/wanted 
    # $WorkingObject

    # Append  current results to final output
    $Output += $WorkingObject

}    

# Script is done looping at this point.  
# We can do something with the $Output.  

# Echo to screen.  REM out if not reqired/wanted 
$Output

# Or output to a file.  The below is an example of going toa  CSV file
# The Output.csv file is located in the same folder as the script.  This is the $PWD or Present Working Directory. 
$Output | Export-Csv -Path $PWD\Output.csv -NoTypeInformation 