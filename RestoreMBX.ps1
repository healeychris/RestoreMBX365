<#	
    .NOTES
    ===========================================================================
    Created with: 	VS
    Created on:   	08/02/2022
    Created by:   	Chris Healey
    Organization: 	
    Filename:     	RestoreMBX.ps1
    Project path:   https://
    Version :       0.1
    ===========================================================================
    .DESCRIPTION
    This script is used to perform a a restore of items from a soft deleted / litigation hold
    mailbox to another online mailbox.
#>
$Version = '0.1'



###### Functions ######
function DisplayExtendedInfo () {

    # Display to notify the operator before running
    Clear-Host
    Write-Host 
    Write-Host 
    Write-Host  '-------------------------------------------------------------------------------'	
	Write-Host  '                   Exchange Online Mailbox Items Restore                       '   -ForegroundColor Green
	Write-Host  '-------------------------------------------------------------------------------'
    Write-Host  '                                                                               '
    Write-Host  '  This Tool kit is used to help restore items from a soft deleted or litigation'    -ForegroundColor YELLOW
    Write-Host  '  on hold mailbox to another cloud mailbox. This script is for cloud           '   -ForegroundColor YELLOW
    Write-Host  "  only operations.                                            version: $version"   -ForegroundColor YELLOW
    Write-Host  '-------------------------------------------------------------------------------'
    Write-Host 
}


# FUNCTION -  Check Active Directory Module is installed
function ImportExchangeOnlineModule () {

    WriteTransactionsLogs -Task "Checking Exchange Online Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        WriteTransactionsLogs -Task "Found Exchange Online Module" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Import-Module ExchangeOnlineManagement	
    } else {
        WriteTransactionsLogs -Task "Failed to locate Exchange Online Module" -Result Error -ErrorMessage "Exchange Module Module not installed" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError false 
        TerminateScript	
    }
    
}


# FUNCTION - Connect to Exchange Online
function ConnectExchangeOnline () {

    try {Get-OrganizationConfig -ea stop | Out-Null;  WriteTransactionsLogs -Task "Existing Exchange Online Connection Found" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        }
    catch {WriteTransactionsLogs -Task "Not Connected to Exchange Online" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        $Connected = $false
        Write-Host `n}
    
    If ($Connected -eq $false){
    try {WriteTransactionsLogs -Task "Connecting to Exchange Online" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false
        Connect-ExchangeOnline -ErrorAction Stop | Out-Null
        Write-Host `n}
    Catch {WriteTransactionsLogs -Task "Unable to Connect to Exchange Online" -Result Error -ErrorMessage "Connect Error" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true
	    Exit
    }
    }
}



# FUNCTION - WriteTransaction Log function    
function WriteTransactionsLogs  {

    #WriteTransactionsLogs -Task 'Creating folder' -Result information  -ScreenMessage true -ShowScreenMessage true exit #Writes to file and screen, basic display
          
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError true #Writes to file and screen and system "error[0]" is recorded
         
    #WriteTransactionsLogs -Task task -Result Error -ErrorMessage errormessage -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError false  #Writes to file and screen but no system "error[0]" is recorded
         


    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$Task,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('Information','Warning','Error','Completed','Processing')]
        [string]$Result,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [string]$ErrorMessage,
    
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateSet('True','False')]
        [string]$ShowScreenMessage,
 
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [string]$ScreenMessageColour,
 
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$IncludeSysError,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$ExportData
)
 
    process {
 
        # Stores Variables
        $LogsFolder           = 'Logs'
 
        # Date
        $DateNow = Get-Date -f g    
        
        # Error Message
        $SysErrorMessage = $error[0].Exception.message

        # Check of log files exist for this session
        If ($Global:TransactionLog -eq $null) {$Global:TransactionLog = ".\TransactionLog_$((get-date).ToString('yyyyMMdd_HHmm')).csv"}
 
        
        # Create Directory Structure
        if (! (Test-Path ".\$LogsFolder")) {new-item -path .\ -name ".\$LogsFolder" -type directory | out-null}
 
  
 
        $TransactionLogScreen = [pscustomobject][ordered]@{}
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Date"-Value $DateNow 
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Task" -Value $Task
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Result" -Value $Result
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "Error" -Value $ErrorMessage
        $TransactionLogScreen | Add-Member -MemberType NoteProperty -Name "SystemError" -Value $SysErrorMessage
        
       
        # Output to screen
       
        if  ($Result -match "Information|Warning" -and $ShowScreenMessage -eq "$true"){
 
        Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
        Write-host " | " -NoNewline
        Write-Host $TransactionLogScreen.Task  -NoNewline
        Write-host " | " -NoNewline
        Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour 
        }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$false"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage  -ForegroundColor $ScreenMessageColour
       }
 
       if  ($Result -eq "Error" -and $ShowScreenMessage -eq "$true" -and $IncludeSysError -eq "$true"){
       Write-host $TransactionLogScreen.Date  -NoNewline -ForegroundColor GREEN
       Write-host " | " -NoNewline
       Write-Host $TransactionLogScreen.Task  -NoNewline
       Write-host " | " -NoNewline
       Write-host $TransactionLogScreen.Result -ForegroundColor $ScreenMessageColour -NoNewline 
       Write-host " | " -NoNewline
       Write-Host $ErrorMessage -NoNewline -ForegroundColor $ScreenMessageColour
       if (!$SysErrorMessage -eq $null) {Write-Host " | " -NoNewline}
       Write-Host $SysErrorMessage -ForegroundColor $ScreenMessageColour
       Write-Host
       }
   
        # Build PScustomObject
        $TransactionLogFile = [pscustomobject][ordered]@{}
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Date"-Value "$datenow"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Task"-Value "$task"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Result"-Value "$result"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "Error"-Value "$ErrorMessage"
        $TransactionLogFile | Add-Member -MemberType NoteProperty -Name "SystemError"-Value "$SysErrorMessage"
 
        # Write Results
        $TransactionLogFile | Export-Csv -Path ".\$LogsFolder\$TransactionLog" -Append -NoTypeInformation
        
        
 
 
        # Clear Error Messages
        $error.clear()
    }   
 
}



# FUNCTION - Ask for Existing Target Mailbox
function TargetMailbox () {
    # Get target mailbox
    Write-Host `n
    $script:TargetMailbox = Read-Host -Prompt "Enter the Target Mailbox using SMTP address: "
    if ($script:TargetMailbox  -eq "") {WriteTransactionsLogs -Task "No Address was entered" -Result Error -ErrorMessage "No ID Entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        TargetMailbox
    }

    Try {$script:TargetMailboxInfo = Get-Mailbox -identity $script:TargetMailbox  -EA Stop

        # Collect specific data from target mailbox
        #$script:TargetMailboxInfoStats | Get-MailboxStatistics | Select-Object ItemCount,TotalItemSize
        $script:TargetMailboxInfoGUID  = $script:TargetMailboxInfo | Select-Object -ExpandProperty ExchangeGUID | Select-Object -ExpandProperty guid       
           
        $Displayname =$script:TargetMailboxInfo.Displayname

        WriteTransactionsLogs -Task "Found Target Mailbox: $Displayname" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False  
        Write-host "======================================================" -ForegroundColor DarkBlue
        Write-host "TARGET :- $script:TargetMailbox | $Displayname | $script:TargetMailboxInfoGUID" 
        Write-host "======================================================" -ForegroundColor DarkBlue
    }
    Catch {WriteTransactionsLogs -Task "Mailbox was not found" -Result Error -ErrorMessage "Not Found in 365" -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError True
        TargetMailbox
    }

    # Confirm Target mailbox is correct
    $ConfirmTarget = Read-Host -Prompt "Are you happy with the selected target mailbox? (Yes/No)"
    if ($ConfirmTarget -eq "Yes") {WriteTransactionsLogs -Task "Target Mailbox Selected $script:TargetMailbox" -Result Error -ErrorMessage "none" -ShowScreenMessage false -ScreenMessageColour GREEN -IncludeSysError False}
    if ($ConfirmTarget -eq "NO") {WriteTransactionsLogs -Task "Target Mailbox Incorrect and asking again" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        TargetMailbox
    }
    if ($ConfirmTarget -notmatch "Yes|No") {WriteTransactionsLogs -Task "No Valid responce entered" -Result Error -ErrorMessage "yes/no not entered" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        TargetMailbox
    }
}

# FUNCTION - Ask for Existing Source Mailbox
function SourceMailbox () {
    # Get Source mailbox
    Write-Host `n
    $script:SourceMailbox = Read-Host -Prompt "Enter the Source Mailbox using SMTP address: "
    if ($script:SourceMailbox  -eq "") {WriteTransactionsLogs -Task "No Address was entered" -Result Error -ErrorMessage "No ID Entered" -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError False
        TargetMailbox
    }

    Try {$script:SourceMailboxInfo = Get-Mailbox -softdeletedmailbox -identity $script:SourceMailbox  -EA Stop

        # Collect specific data from Source mailbox
        #$script:SourceMailboxInfoStats | Get-MailboxStatistics | Select-Object ItemCount,TotalItemSize
        $script:SourceMailboxInfoGUID  = $script:SourceMailboxInfo | Select-Object -ExpandProperty ExchangeGUID | Select-Object -ExpandProperty guid       
           
        $Displayname =$script:SourceMailboxInfo.Displayname

        WriteTransactionsLogs -Task "Found Source Mailbox: $Displayname" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError False  
        Write-host "======================================================" -ForegroundColor DarkBlue
        Write-host "SOURCE:- $script:SourceMailbox | $Displayname | $script:SourceMailboxInfoGUID" 
        Write-host "======================================================" -ForegroundColor DarkBlue
    }
    Catch {WriteTransactionsLogs -Task "Mailbox was not found" -Result Error -ErrorMessage "Not Found in 365" -ShowScreenMessage true -ScreenMessageColour red -IncludeSysError True
        SourceMailbox
    }

    # Confirm Source mailbox is correct
    $ConfirmSource = Read-Host -Prompt "Are you happy with the selected Source mailbox? (Yes/No)"
    if ($ConfirmSource -eq "Yes") {WriteTransactionsLogs -Task "Source Mailbox Selected $SourceMailbox" -Result Information -ErrorMessage "none" -ShowScreenMessage false -ScreenMessageColour GREEN -IncludeSysError False}
    if ($ConfirmSource -eq "NO") {WriteTransactionsLogs -Task "Source Mailbox Incorrect and asking again" -Result Information -ErrorMessage "none" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        TargetMailbox
    }
    if ($ConfirmSource -notmatch "Yes|No") {WriteTransactionsLogs -Task "No Valid responce entered" -Result Error -ErrorMessage "yes/no not entered" -ShowScreenMessage true -ScreenMessageColour YELLOW -IncludeSysError False
        SourceMailbox
    }
}


# FUNCTION -  Perform Restore of mailbox data
function PerformRestore () {

    
    
    try {New-MailboxRestoreRequest -SourceMailbox $script:SourceMailboxInfoGUID -TargetMailbox $script:TargetMailboxInfoGUID -Name $script:TargetMailbox -AllowLegacyDNMismatch  -EA STOP
        WriteTransactionsLogs -Task "Performing Mail Restore from $script:SourceMailbox -> $script:TargetMailbox" -Result Information -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour GREEN -IncludeSysError false 
        
        Write-Host
        Write-host "To check the progress use Powershell and Get-MailboxRestoreRequest -Name $TargetMailbox | fl"
        }
    Catch {WriteTransactionsLogs -Task "Performing Mail Restore from $script:SourceMailbox -> $script:TargetMailbox Failed!!!" -Result "Error" -ErrorMessage none -ShowScreenMessage true -ScreenMessageColour RED -IncludeSysError true }
    
}



# Run Order
DisplayExtendedInfo
ImportExchangeOnlineModule
ConnectExchangeOnline
SourceMailbox
TargetMailbox
PerformRestore



