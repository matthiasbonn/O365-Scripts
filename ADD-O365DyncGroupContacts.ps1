#requires -version 2
<#
.SYNOPSIS
    Create or update an existing Mail Contant in Office 365 for dynamic distribution groups in an on prem Exchange infrastructure.

.DESCRIPTION
    To address the dynamic distribution groups from the on premis exchange a coresponending mail contanct in office 365 is needed. 
    The script fetsch all dynamic distribution groups opbjects and stores the following attributs in a SQLIte DB,

        * DisplayName 
        * Alias 
        * PrimarySMTPAddress
        * WhenChanged
        * GUID

    If a dynamic distribution group is deleted in exchange on premis the record in the database is flaged to delete the correspondenting 
    MailContact in Office 365. 

    The scripts updates the db rord if the WhenChange Timestap has a diffenrece greater or equal one minute, there fore the mailcontact is also 
    updated if the WhenChange Attribut in the database is newer the the attribut on the mailcontact in office 365



.PARAMETER SQLiteDB
   Path to the SQLite Database File to store the information

.INPUTS
    None

.OUTPUTS
  Logfile in a Sub-Dirctory Logs under the script path.

.NOTES
  Version:        1.0
  Author:         Bonn, Matthias - Alegri International Service GmbH
  Creation Date:  03.02.2016
  Purpose/Change: Initial script development
  
.EXAMPLE
    ADD-O365DynGroupContatcs.ps1
#> 

[CmdletBinding(SupportsShouldProcess=$True)]
param (
             [Parameter(Mandatory=$False)]
             [ValidateNotNullorEmpty()]
             [string]$SQLiteDB = 'E:\MSX_Scripts\SyncDynGroups.SQLite'
       )
#region Initialization code
    Write-Verbose 'Initialize stuff in Begin block'
    #region Modules / Addins / DotSourcing
    
    $LoadedModules = Get-Module  -name PSSQLite | Select Name
    If (!$LoadedModules -eq 'PSSQLite') {
        Try{
            Import-Module PSSQLite
        }
        Catch{
            Write-Host 'Module PSSQLite does not exists'
        }
    } 

    #if (-not(Get-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction Silentlycontinue)){Add-PSSnapin Quest.ActiveRoles.ADManagement}

    #endregion Modules / Addins / DotSourcing
    
    #Dump all existing Variables in a variable for housekeeping
    $startupVariables=''
    new-variable -force -name startupVariables -value ( Get-Variable | % { $_.Name } )

    $ScriptVersion   = 0.1
    $ScriptPath      = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
    $ScriptName     = [system.io.path]::GetFileNameWithoutExtension($MyInvocation.InvocationName)
    If (!$Log){
        $LogPath     = "$ScriptPath\Logs"
        }
    Else
    {
        $LogPath     = $Log
    }
    $LogFile         = "$Logpath\$Scriptname.log"
    $DateFormat      = Get-Date -Format 'yyyyMMdd_HHmmss'
    Write-Verbose "Start Script Version $ScriptVersion at $DateFormat from $ScriptPath" -verbose
    IF(!(Test-Path $LogPath)) {mkdir $LogPath}
    
    $i=0
    $ObjectCount=0

    #region nlog

        # configuration for the nlog feature
        [Reflection.Assembly]::LoadFile(“$scriptPath\NLog.dll”)
        $nlogconfig = new-object NLog.Config.XmlLoggingConfiguration("$ScriptPath\NLog.config")
        ([NLog.LogManager]::Configuration)=$nlogconfig
        [NLog.Targets.FileTarget]$fileTarget = [NLog.Targets.FileTarget]([NLog.LogManager]::Configuration.FindTargetByName("logfile"))
        $fileTarget.FileName = $LogFile
        $PSlogger = [NLog.LogManager]::GetLogger('PSLogger')

    #endregion nlog 

#endregion Initialization code


#region functions

function Check-Powershell64{
$is64Bit=[Environment]::Is64BitProcess
return $is64Bit
}

function Cleanup-Variables {

  Get-Variable |

    Where-Object { $startupVariables -notcontains $_.Name } |

    % { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" }

}

Function Test-Credential { 
    [OutputType([Bool])] 
     
    Param ( 
        [Parameter( 
            Mandatory = $true, 
            ValueFromPipeLine = $true, 
            ValueFromPipelineByPropertyName = $true 
        )] 
        [Alias( 
            'PSCredential' 
        )] 
        [ValidateNotNull()] 
        [System.Management.Automation.PSCredential] 
        [System.Management.Automation.Credential()] 
        $Credential = [System.Management.Automation.PSCredential]::Empty, 
 
        [Parameter()] 
        [String] 
        $Domain = $env:USERDOMAIN 
    ) 
 
    Begin { 
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") 
        $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain, $Domain) 
    } 
 
    Process { 
        $networkCredetial = $Credential.GetNetworkCredential() 
        return $principalContext.ValidateCredentials($networkCredetial.UserName, $networkCredetial.Password) 
    } 
 
    End { 
        $principalContext.Dispose() 
    } 
} 


#endregion fuctions


#region Process data
$PSlogger.Info('Begin of execution')
IF(Check-Powershell64){$PSlogger.Info('Running in 64 BIT Process..')}
do{
    $O365_User= Get-Credential -Message 'Please enter O365 User Information'
}
Until (Test-Credential $O365_User)

Try{
    # Connect to O365 an use a prefix so that boath enviroments can be manageded
    $o365Session= New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $O365_User -Authentication Basic -AllowRedirection
    Import-PSSession $o365Session -Prefix O365
    }
Catch{
    $PSlogger.Error("Can´t connect to Office 365 inftrasucture")
    Exit (1)
}
Try{
    $ExSession= New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://HBMES554.hugoboss.com/PowerShell/ -Authentication Kerberos
    Import-PSSession $ExSession
    }
Catch{
    $PSlogger.Error("Can´t connect to Exchange on prem inftrasucture")
    Exit (1)

}

Try {
    $SQLiteDBCon = New-SQLiteConnection -DataSource $SQLiteDB
    $PSlogger.Info("Open SQLite DB $SQLiteDB")
    }
Catch{
    $PSlogger.Error("Can´t open SQLite DB $SQLiteDB")
    Exit (1)
}

#region ActiveDirctory-2-DB
$PSlogger.Info('Fetch all dynamich Distribution Groups')
$Groups = Get-DynamicDistributionGroup -ResultSize Unlimited

# Set DeleteStatus to all records to identify deleted dynamic distribution groups in the on prem infrastructure
Invoke-SqliteQuery -DataSource $SQLiteDB -Query "Update SYNC_DNYGROUPSOPBJECT Set DeleteStatus='deleting' Where DeleteStatus=''"

Foreach ($group in $groups){
    $DBRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "Select * from SYNC_DNYGROUPSOPBJECT Where PrimarySMTPAddress=""$($group.PrimarySmtpAddress)"""
    If ($DBRecord.PrimarySMTPAddress){
       If ($($group.WhenChanged - [datetime]$DBRecord.WhenChanced).Minutes -eq 0){
            # Reset the deleteStatus for existing dynamic distribution groups in the database
            $PSlogger.Info("No Changes in DB| $($group.PrimarySMTPAddress)")
            $query = "Update SYNC_DNYGROUPSOPBJECT Set DeleteStatus='' Where PrimarySMTPAddress=""$($group.PrimarySMTPAddress)"""
       }
       Else
       {
            # Update of the databaserecord the timedifferenc is greater then one minute for extisting objects
            $PSlogger.Info("Update in DB | $($group.PrimarySMTPAddress)")
            $query = "Update SYNC_DNYGROUPSOPBJECT set PrimarySMTPAddress=""$($group.PrimarySmtpAddress)"", DisplayName=""$($group.DisplayName)"", Alias=""$($group.Alias)"", DeleteStatus='',WhenChanced=""$($group.WhenChanged)"" Where PrimarySMTPAddress=""$($group.PrimarySMTPAddress)"""
       }
    }
    Else
    {
        # Inserst database record for new objects
        $PSlogger.Info("Insert in DB | $($group.PrimarySMTPAddress)")
        $query = "Insert into SYNC_DNYGROUPSOPBJECT (PrimarySMTPAddress,DisplayName,Alias,DeleteStatus,GUID,WhenCreated,WhenChanced) Values (""$($group.PrimarySmtpAddress)"", ""$($group.DisplayName)"", ""$($group.Alias)"",'',""$($group.GUID)"", ""$($group.WhenCreated)"", ""$($group.WhenChanged)"")"
    }
    
    Invoke-SqliteQuery -DataSource $SQLiteDB -Query $query
        
}


$DBRecord = ''
#endregion ActiveDirctory-2-DB

#region update_O365Tenant

# detect all db records for active objects
$DBRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query 'Select * from SYNC_DNYGROUPSOPBJECT Where DeleteStatus=""'
For ($i=0; $i -le $DBRecord.Count -1; $i++){
    # Try to get the corresondenting mailcontact for the distribution group in office 365
    $O365Contact = Get-O365MailContact -Identity $($DBRecord[$i].PrimarySmtpAddress)

    IF(!$O365Contact){
        # No contact exists in the O365 environment - create new object   
        New-O365MailContact -DisplayName $($DBRecord[$i].DisplayName) -Name $($DBRecord[$i].DisplayName) -ExternalEmailAddress $($DBRecord[$i].PrimarySmtpAddress) -Alias $($DBRecord[$i].Alias)
        $PSlogger.Info("Create Contact in O365 for for $($DBRecord[$i].PrimarySmtpAddress)")
        $O365Contact = Get-O365MailContact -Identity $($DBRecord[$i].PrimarySmtpAddress)
        $Query = "Update SYNC_DNYGROUPSOPBJECT set O365_WhenChaned=""$($O365Contact.WhenChanged)"", O365_Guid=""$($O365Contact.Guid)"" Where PrimarySMTPAddress=""$($O365Contact.PrimarySMTPAddress)"""
        Invoke-SqliteQuery -DataSource $SQLiteDB -Query $query 
    }
    ElseIf ($O365Contact.WhenChanged  -ge [Datetime]$DBRecord[$i].WhenChanced){
        # No changes are needed because the O365 Object has a newer WhenChanced Date
        $PSlogger.Info("No Change need in Office 3365 for $($DBRecord[$i].PrimarySmtpAddress)")
    }
    Else
    {
        #Update of the o365 needed 
        $PSlogger.Info("Update for $($DBRecord[$i].PrimarySmtpAddress)")
        Set-O365MailContact -Identity $($DBRecord[$i].PrimarySmtpAddress) -DisplayName $($DBRecord[$i].DisplayName) -Name $($DBRecord[$i].DisplayName) -Alias $($DBRecord[$i].Alias)
        # Writeback some information about the o365 obtect to the database record
        $O365Contact = Get-O365MailContact -Identity $($DBRecord[$i].PrimarySmtpAddress)
        $Query = "Update SYNC_DNYGROUPSOPBJECT set O365_WhenChanged=""$($O365Contact.WhenChanged)"" Where PrimarySMTPAddress=""$($O365Contact.PrimarySMTPAddress)"""
        Invoke-SqliteQuery -DataSource $SQLiteDB -Query $query
    }
}

# Get all records that shoud be deleted in o365
$DBRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query 'Select * from SYNC_DNYGROUPSOPBJECT Where DeleteStatus="deleting"'
For ($i=0; $i -le $DBRecord.Count -1; $i++){
    # get the infmation from o354 an remove the mail contact
    $O365Contact = Get-O365MailContact -Identity $($DBRecord[$i].PrimarySmtpAddress)
    Remove-O365MailContact -Identity $DBRecord.PrimarySMTPAddress -Confirm:$False
    $PSlogger.Info("Contact in O365 removed for $DBRecord[$i].PrimarySmtpAddress")    
    # Write back some information to the database recored, e.q. set DeleteSatus to deleted
    $Query = "Update SYNC_DNYGROUPSOPBJECT set O365_WhenChanged=""$($O365Contact.WhenChanged)"",DeleteStatus='deleted' Where PrimarySMTPAddress=""$($O365Contact.PrimarySMTPAddress)"""
    Invoke-SqliteQuery -DataSource $SQLiteDB -Query $query
}

#endregion update_O365Tenant

#endregion Process data

#region Finalizing 
    Remove-PSSession $o365Session
    Remove-PSSession $ExSession
    $PSlogger.Info('Cleanup started ...')
    Cleanup-Variables
#endregion Finalizing  