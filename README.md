# Script To Automate Exchange Health Manager Service For Monitoring Mailbox Reset
 Script To Automate Exchange Health Manager Service For Monitoring Mailbox Reset


.SYNOPSIS
Purpose of this script to help Exchange administrators with the process of resetting Exchange 2013 and 2016 Health Mailboxes.
For more details, please see this post:

https://blogs.technet.microsoft.com/rmilne/2016/06/16/script-to-automate-exchange-health-manager-service-for-monitoring-mailbox-reset/

  

There is logic to loop through multiple Exchange servers. A filter will be used to exclude non Exchange 2013 servers in the default configuration. 

   

Filters can be adjusted to suit the particular task that is required.  There are multiple    filtering examples here:

                http://blogs.technet.com/b/rmilne/archive/2014/03/24/powershell-filtering-examples.aspx

 

This is the line that you would need to change:
  $ExchangeServers = Get-ExchangeServer $ServerString | Where-Object {$_.AdminDisplayVersion -match "^Version 15"}

  Version 2 added an additional qualifier so that if you had a consistent naming strategy, you could select servers from a particular site using a wildcard. 
Format to use would be:
    ExchYYZ* 

 If you want to select all servers use *



This is not considered normal maintenance and thus did not want to automate end to end, else someone may be tempted to schedule this as a task every Sunday morning

One script does not have the knowledge to deal with every customer environment.  Thus you will still need to do the AD work etc, and confirm the results in your environment.

 

Refer to the blog post for all the details.               

Script to be executed on an Exchange 2013 or 2016 server

 

 

.DESCRIPTION
Script will use WMI object to work with remote windows services.  Get-Service, Set-Service, Stop-Service and Start-Service can be used for this, though WMI was used here.

As an example, the below returns a reference to the service on a machine called DC1

Service = Get-WmiObject -ComputerName DC1 -Class Win32_Service -Filter "Name='MSExchangeHM'"

Script requires a command line parameter.  You must specify the start/stop action to take, though in current version of the script it will default to check. 

Accepted actions are:

                                Start

                                Stop

                                Check

Stop  -- will stop the Microsoft Exchange Health Manager service (MSExchangeHM) on the servers returned in the collection

Start -- will stop the Microsoft Exchange Health Manager service (MSExchangeHM) on the servers returned in the collection

Check -- will display the status of the Microsoft Exchange Health Manager service (MSExchangeHM) on the servers returned in the collection

 

CmdletBinding was added to check for the command line paramater being added, and to check the values entered.

It also allows you to tab complete them, as an added bonus.  Sweet!

 

.ASSUMPTIONS
Script is being executed with sufficient permissions to access the server(s) targeted.

Script is executed in the Exchange Management Shell

Script is executed on Exchange 2013 or 2016 server

Script is for Exchange 2013 and Exchange 2016

Thus you have the correct version of PowerShell and WMF installed

You can live with the Write-Host cmdlets :)

You can add your error handling if you need it. 

 

.VERSION
1.0  12-6-2016 -- Initial script

1.1  16-6-2016 -- Initial script released to the scripting gallery

1.2  4-10-2016 -- Updated with Feedback from Feras

  