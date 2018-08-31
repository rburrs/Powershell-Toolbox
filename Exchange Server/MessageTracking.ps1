<#
.SYNOPSIS
    Get Exchange message logs

.DESCRIPTION
    Function to get messages sent by a particular sender during a certain timeframe.  The script first 
    requests credentials for an account with permissions to EXchange and establishes a remote session to 
    the Exchange CAS.  The output is exported to a CSV file 
    that is save to the logged in user's desktop.

.PARAMETER Sender
    The SMTP address of the message sender

.PARAMETER Start
    The start date and time

.PARAMETER  End
    The end date and time

.PARAMETER Subject
    The subject of the message

.EXAMPLE
    Get-MessageTrackingLog -Sender "jane.doe@alston.com" -Start "01/31/2017 05:15:10" -End "02/01/2017 13:45:00" 

.NOTES
    This function requires admininstrator priviledges and Exchange Powershell cmdlets
#>

Function Get-Messages {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true,
            Position=0,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="Enter the SMTP address of the sender.")]
            [AllowEmptyString()]
            [AllowNull()]
            [string]$Sender,
        [parameter(Mandatory=$true,
            Position=1,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="Enter the SMTP address of the recipient.")]
            [AllowEmptyString()]
            [AllowNull()]
            [string]$Recipient,
        [parameter(Mandatory=$true,
            Position=2,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="Enter start date as 04/20/2017 01:00:00.")]
            [datetime]$Start, 
        [parameter(Mandatory=$true,
            Position=3,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="Enter end date as 04/20/2017 18:00:00.")]
            [datetime]$End,
        [parameter(
            Position=4,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)]
            [AllowEmptyString()]
            [AllowNull()]
            [string]$Subject
        
          )

    
$servers = Get-TransportServer 
$filepath = "$($env:USERPROFILE)\desktop\messages.csv"
        
         
    foreach ($server in $servers) {  

        $servers.name | Get-MessageTrackingLog -Sender $Sender -Start $Start -End $End | 
        Select-Object TimeStamp,Sender,@{name="Recipients";Expression={$_.Recipients}},MessageSubject | 
        Export-csv $Filepath -NoTypeInformation

    }
  
}

Get-Messages

