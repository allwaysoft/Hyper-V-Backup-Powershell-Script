#region
##############################################################################################
# Script zum Export von VMs unter Hyper-V mit Windows Server 2016 oder Windows 10 Pro        #
# erstellt von Jan Kappen - j.kappen@rachfahl.de                                             #
# Version 0.4.1                                                                              #
# 06. Mai 2017                                                                               #
# Diese Script wird bereitgestellt wie es ist, ohne jegliche Garantie. Der Einsatz           #
# erfolgt auf eigene Gefahr. Es wird jegliche Haftung ausgeschlossen.                        #
# Wer dem Autor etwas Gutes tun möchte, er trinkt gern ein kaltes Corona :)                  #
#                                                                                            #
# Dieses Script beinhaltet keine Hilfe. RTFM.                                                #
#                                                                                            #
# www.hyper-v-server.de | www.rachfahl.de                                                    #
#                                                                                            #
##############################################################################################
#endregion

Param(
	[string] $VM ="",
	[string] $exportpath = "",
	[string] $logpath = "",
	[switch] $save,
    [switch] $ProductionCheckpoint,
    [switch] $shutdown,
    [switch] $statusmail
)

function Mail
{
    Write-Host -ForegroundColor Red (Get-Date) "The mail configuration has to be adjusted, otherwise no mailing is possible!"
    # $mail = @{
    #    SmtpServer = 'mailer.domain.loc'
    #    Port = 25
    #    From = 'backupskript@rachfahl.de'
    #    To = 'empfaenger@domain.loc'
    #    Subject = "'Script backup of THE VM' $VM"
    #    Body = "'Here is the log of the export backup of VM' $VM"
    #    Attachments = "$logfile"
    # }
    # Send-MailMessage @mail
}

#region logfile
$logfileDate = Get-Date -Format yyyy-MM-dd
if (!$logpath) {
                    $logPathPresent = Test-Path ${env:homedrive}\windows\Logs\HyperVExport\
                    if ($logPathPresent -eq $False) { new-item ${env:homedrive}\windows\Logs\HyperVExport\ -itemtype directory }
                    $logfile = "${env:homedrive}\windows\Logs\HyperVExport\$logfileDate.log" 
                }
                else 
                { $logfile = "$logpath\$logfileDate.log" }
# Start logging
Start-Transcript -Path $logfile -Append
#endregion

#region Abfrage auf benötigte Parameter
if (!$VM) { 
              Write-Host -ForegroundColor Red (Get-Date) "Parameter -VM must exist and contain a name. Abort!"
              exit 
          }
if (!$exportpath) { 
                      Write-Host -ForegroundColor Red (Get-Date)  "Parameters -RemotePath must exist and contain a path. Abort!"
                      exit
                  }
#endregion

#region Export der VM
if ($save -match "true")
{
    Write-Host -ForegroundColor Green (Get-Date) "VM is saved"
    Save-VM -Name $VM -verbose
    Write-Host -ForegroundColor Green (Get-Date) "Exporting the VM"
    $DestinationPathPresent = Test-Path $exportpath\$VM
    if 
        ($DestinationPathPresent -eq $False) { Export-VM -Name $VM -Path $exportpath -verbose }
    else 
        { Remove-Item -Recurse -Force $exportpath\$VM -verbose; Export-VM -Name $VM -Path $exportpath -verbose }
    Write-Host -ForegroundColor Green (Get-Date) "Export completes, VM is turned back on"
    Start-VM -Name $VM -verbose
    if ($statusmail -match "true")
        { Mail; exit }
    else
        { exit }
}

if ($shutdown -match "true")
{
    $ShutdownStatus = get-vm -Name $VM | Get-VMIntegrationService | where { $_.Name -EQ "Shutdown" -or $_.Name -EQ "shutdown" }
    if ($ShutdownStatus.Enabled -eq "True")
    {
        Write-Host -ForegroundColor Green (Get-Date) "Export seems possible, VM shuts down"
        Stop-VM -Name $VM -Force -verbose
        Write-Host -ForegroundColor Green (Get-Date) "Exporting the VM"
        $DestinationPathPresent = Test-Path $exportpath\$VM
        if 
            ($DestinationPathPresent -eq $False) { Export-VM -Name $VM -Path $exportpath -verbose }
        else 
            { Remove-Item -Recurse -Force $exportpath\$VM -verbose; Export-VM -Name $VM -Path $exportpath -verbose }
        Write-Host -ForegroundColor Green (Get-Date) "VM is turned back on"
        Start-VM -Name $VM -verbose
        if ($statusmail -match "true")
            { Mail; exit }
        else
            { exit }
    }
    else
    {
        Write-Host -ForegroundColor Red (Get-Date) "Export does not seem to be possible, operation is canceled!"
        exit
    }
}

if ($ProductionCheckpoint -match "true")
{
    $SnapshotName = "ExportScriptCheckpoint"
    Write-Host -ForegroundColor Green (Get-Date) "Checkpoint is Create"
    $DestinationPathPresent = Test-Path $exportpath\$VM
    if 
        ($DestinationPathPresent -eq $False) 
        { Checkpoint-VM -Name $VM -SnapshotName $SnapshotName -verbose
          Export-VMSnapshot -VMName $VM -Name $SnapshotName -Path $exportpath -verbose
          Remove-VMSnapshot -VMName $VM -Name $SnapshotName -verbose
        }
    else 
        { Remove-Item -Recurse -Force $exportpath\$VM -verbose
          Checkpoint-VM -Name $VM -SnapshotName $SnapshotName -verbose
          Export-VMSnapshot -VMName $VM -Name $SnapshotName -Path $exportpath -verbose
          Remove-VMSnapshot -VMName $VM -Name $SnapshotName -verbose
        }
    Write-Host -ForegroundColor Green (Get-Date) "Export completed"
    if ($statusmail -match "true")
        { Mail; exit }
    else
        { exit }
}
else
{
    Write-Host -ForegroundColor Green (Get-Date) "VM is exported online"
    Write-Host -ForegroundColor Green (Get-Date) "Exporting the VM"
    $DestinationPathPresent = Test-Path $exportpath\$VM -verbose
    if 
        ($DestinationPathPresent -eq $False) { Export-VM -Name $VM -Path $exportpath -verbose }
    else 
        { Remove-Item -Recurse -Force $exportpath\$VM -verbose; Export-VM -Name $VM -Path $exportpath -verbose }
    Write-Host -ForegroundColor Green (Get-Date) "Export of VM $VM completed"
    if ($statusmail -match "true")
        { Mail; exit }
    else
        { exit }
}
#endregion