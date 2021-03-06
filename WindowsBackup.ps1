#Pre-Reqs:
#All Computers:
    #You must run "Winrm quickconfig" on each computer you want to remotly administer with powershell
    
#Server 2008 R2
    #You must install the "Windows Server Backup" Feature with command line support

#Windows 7
    #Professional Version or above

#////////////////////////////////////
#               WSB
#///////////////////////////////////

Function Is-WSBInstalled{
    param($RemoteToCompName)

    $RemoteSession = new-pssession -computername $RemoteToCompName
    
    If ($RemoteSession.Availability -eq "Available") {
        $RemoteWSBSnapIn = invoke-command -session $RemoteSession -scriptblock {
            add-pssnapin Windows.serverbackup -ErrorAction SilentlyContinue
            get-pssnapin Windows.serverbackup -ErrorAction SilentlyContinue
        }
    }
    
    Remove-PSSession -Session $RemoteSession

    If($RemoteWSBSnapIn -ne $null){
        Return $true
    }
    Else{
        Return $false
    }
}


Function Run-WSBBackup {
param ( [OBJECT] $BackupJob )

    $BackupReturn = New-Object Object
    $BackupReturn | add-member NoteProperty Status ''
    $BackupReturn | add-member NoteProperty DestSubFolder ''
    $BackupReturn | add-member NoteProperty DestFullPath ''
    $BackupReturn | add-member NoteProperty PreviousBackupSize (New-Object Object)
    $BackupReturn | add-member NoteProperty BackupSize (New-Object Object)
    $BackupReturn | add-member NoteProperty CommandOutput (New-Object System.Collections.ArrayList)
    $BackupReturn | add-member NoteProperty ErrorOutput ''
    
    $BackupReturn.DestSubFolder = "WindowsImageBackup\$($BackupJob.TargetServer)"
    $BackupReturn.DestFullPath = "$($BackupJob.DestFileShare)\$($BackupReturn.DestSubFolder)"

    #Take the last version and keep it in case the backup fails
    If($BackupJob.UsePriorBackupIfFailure -eq $true){
        If(test-path $($BackupReturn.DestFullPath)){
            Rename-Item "$($BackupReturn.DestFullPath)" "$($BackupReturn.DestFullPath).History"
            $BackupReturn.PreviousBackupSize = (Get-DiskUsed "$($BackupReturn.DestFullPath).History")
        }
    }
    
    $RemoteSession = new-pssession -computername $BackupJob.TargetServer
    
    If ($RemoteSession.Availability -eq "Available") {
        $BackupReturn.CommandOutput = invoke-command -session $RemoteSession -scriptblock {
        param( [OBJECT] $BackupJob )
                
				#This is a bug workaround where if the destination server is not in the arp cache, the backup would fail.
                nbtstat -a "$((($BackupJob.DestFileShare).Split('\'))[2])"
                Add-Pssnapin Windows.serverbackup
                
                $WBPolicy = New-WBPolicy
                
                If($BackupJob.BareMetalRecovery -eq 'YES'){
                    Add-WBBareMetalRecovery -Policy $WBPolicy
                }

                #Create the credential to Run the Backup
                $PasswordSecureString = ConvertTo-SecureString $BackupJob.Password -AsPlainText -Force
                $BackupJobCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($BackupJob.UserName),$PasswordSecureString
                $WBTargetPolicy = New-WBBackupTarget -NetworkPath $BackupJob.DestFileShare -Credential $BackupJobCredential

                Add-WBBackupTarget -Policy $WBPolicy -Target $WBTargetPolicy

                #Set VSS Options
                If($BackupJob.VSSOptions -eq 'COPY'){
                    Set-WBVssBackupOptions -Policy $WBPolicy -VSSCopyBackup
                }
                Else {
                    Set-WBVssBackupOptions -Policy $WBPolicy -VSSFullBackup
                }
                    
                #Add additional volumes
                
                If($BackupJob.AdditionalVolumes -ne ''){
                    Foreach($AdditionalVolume in ($BackupJob.AdditionalVolumes).Split("|")){
                        $WBVolume = Get-WBVolume -VolumePath $AdditionalVolume
                        Add-WBVolume -Policy $WBPolicy -Volume $WBVolume
                    }
                }
                
                #Initiate the Backup Policy
                Start-WBBackup $WBPolicy
                
        } -ArgumentList $BackupJob
    }
    
    Remove-PSSession -Session $RemoteSession

    If(Is-WSBBackupSuccess $BackupJob.TargetServer){
        $BackupReturn.Status = 'SUCCESS'
        If($BackupJob.UsePriorBackupIfFailure -eq $true){
            If(test-path "$($BackupReturn.DestFullPath).History"){
                Remove-Item -path "$($BackupReturn.DestFullPath).History" -Force -Recurse
            }
        }
    }
    Else{
        $BackupReturn.Status = 'FAILURE'
        If($BackupJob.UsePriorBackupIfFailure -eq $true){
            If(test-path $BackupReturn.DestFullPath){
                Remove-Item -path "$($BackupReturn.DestFullPath)" -Force -Recurse
            }
            If(test-path "$($BackupReturn.DestFullPath).History"){
                Rename-Item "$($BackupReturn.DestFullPath).History" "$($BackupReturn.DestFullPath)"
            }
        }

        $BackupReturn.ErrorOutput = Get-WSBBackupErrorMessage $BackupJob.TargetServer
    }

    $BackupReturn.BackupSize = (Get-DiskUsed "$($BackupReturn.DestFullPath)")

    Return $BackupReturn
}


Function Is-WSBBackupSuccess {
    param( [STRING] $RemoteToCompName,
            [INT] $TryCount = 0)
    
    $BackupResult = New-Object Object
    $BackupResult | add-member NoteProperty Status ''
    $BackupResult | add-member NoteProperty ErrorMessage @()

    $RemoteSession = new-pssession -computername $RemoteToCompName
    
    If ($RemoteSession.Availability -eq "Available") {
        $WSBErrorDesc = invoke-command -session $RemoteSession -scriptblock {
            Add-Pssnapin Windows.serverbackup
            $WSBJob = get-wbjob -previous 1
            $WSBJob.ErrorDescription
        }
    }
    Else{#The computer is not available, try to get result again.
        If($TryCount -lt 2){
            $TryCount += 1
            Start-Sleep -Seconds 10
            Return (Is-WSBBackupSuccess $RemoteToCompName $TryCount)
        }
    }
    
    Remove-PSSession -Session $RemoteSession

    If($WSBErrorDesc -eq ''){
        Return $true
    }
    Else {
        Return $false
    }
}


Function Get-WSBBackupErrorMessage {
    param($RemoteToCompName)

    $RemoteSession = new-pssession -computername $RemoteToCompName
    
    If ($RemoteSession.Availability -eq "Available") {
        [System.Collections.ArrayList] $WSBErrorDesc = invoke-command -session $RemoteSession -scriptblock {
            Add-Pssnapin Windows.serverbackup
            $WSBJob = get-wbjob -previous 1
            $WSBJob.ErrorDescription
        }
    }
    Else{#The computer is not available, try to get result again.
        Start-Sleep -Seconds 10
        Return (Is-WSBBackupSuccess $RemoteToCompName)
    }
    
    Remove-PSSession -Session $RemoteSession

    Return $WSBErrorDesc
}


#////////////////////////////////////
#               wbadmin
#///////////////////////////////////

Function Run-WBAdminBackup {
 param ([OBJECT] $BackupJob)
    
    $BackupReturn = New-Object Object
    $BackupReturn | add-member NoteProperty Status ''
    $BackupReturn | add-member NoteProperty DestSubFolder ''
    $BackupReturn | add-member NoteProperty DestFullPath ''
    $BackupReturn | add-member NoteProperty PreviousBackupSize (New-Object Object)
    $BackupReturn | add-member NoteProperty BackupSize (New-Object Object)
    $BackupReturn | add-member NoteProperty CommandOutput (New-Object System.Collections.ArrayList)
    $BackupReturn | add-member NoteProperty ErrorOutput (New-Object System.Collections.ArrayList)

    $BackupReturn.DestSubFolder = "WindowsImageBackup\$($BackupJob.TargetServer)"
    $BackupReturn.DestFullPath = "$($BackupJob.DestFileShare)\$($BackupReturn.DestSubFolder)"

    #Take the last version and keep it in case the backup fails
    If($BackupJob.UsePriorBackupIfFailure -eq $true){
        If(test-path $($BackupReturn.DestFullPath)){
            Rename-Item "$($BackupReturn.DestFullPath)" "$($BackupReturn.DestFullPath).History"
            $BackupReturn.PreviousBackupSize = (Get-DiskUsed "$($BackupReturn.DestFullPath).History")
        }
    }
    
    [DateTime] $BackupStartUTC = (Get-Date).ToUniversalTime()
    $RemoteSession = new-pssession -computername $BackupJob.TargetServer

    If ($RemoteSession.Availability -eq "Available") {

        $BackupReturn.CommandOutput = invoke-command -session $RemoteSession -scriptblock {
            param([OBJECT] $BackupJob)
             
                #build
                $WBAdminCommand = "wbadmin start backup"
                $WBAdminCommand += " -backupTarget:""$($BackupJob.DestFileShare)""" #The above extra information is appended to this file path
                $WBAdminCommand += " -user:""$($BackupJob.UserName)"""
                $WBAdminCommand += " -password:""$($BackupJob.Password)"""
                If($BackupJob.BareMetalRecovery -eq 'YES'){$WBAdminCommand += " -allCritical"}
                If($BackupJob.VSSOptions -eq 'COPY') { $WBAdminCommand += " -vssCopy" } Else{ $WBAdminCommand += " -vssFull" }
                If($BackupJob.AdditionalVolumes -ne ''){
                    $WBAdminCommand += " -include:$(($BackupJob.AdditionalVolumes).Replace('|',','))"
                }
                $WBAdminCommand += " -quiet"   #Start the backup without confirmation

                #The below command is a bug fix/workaround for Windows Backup.  If the destination file share computer name is not in the ARP table, the backups do not consistently run properly.
                nbtstat -a "$((($BackupJob.DestFileShare).Split('\'))[2])"
                
                #Execute the Backup Job
                Invoke-Expression "& $WBAdminCommand"
             
        } -ArgumentList $BackupJob
    }

    Remove-PSSession -Session $RemoteSession

    If(Is-WBAdminBackupSuccess $BackupJob.TargetServer $BackupStartUTC){
        $BackupReturn.Status = 'SUCCESS'
        If($BackupJob.UsePriorBackupIfFailure -eq $true){
            If(test-path "$($BackupReturn.DestFullPath).History"){
                Remove-Item -path "$($BackupReturn.DestFullPath).History" -Force -Recurse
            }
        }
    }
    Else{
        $BackupReturn.Status = 'FAILURE'
        If($BackupJob.UsePriorBackupIfFailure -eq $true){
            If(test-path $BackupReturn.DestFullPath){
                Remove-Item -path "$($BackupReturn.DestFullPath)" -Force -Recurse
            }
            If(test-path "$($BackupReturn.DestFullPath).History"){
                Rename-Item "$($BackupReturn.DestFullPath).History" "$($BackupReturn.DestFullPath)"
            }
        }
    }
    
    $BackupReturn.BackupSize = (Get-DiskUsed "$($BackupReturn.DestFullPath)")

    Return $BackupReturn
}

Function Is-WBAdminBackupSuccess{
    param([STRING]$RemoteToCompName,
            [DateTime]$BackupStartDate,
            [INT] $TryCount = 0)
    
    $RemoteSession = new-pssession -computername $RemoteToCompName
    
    If ($RemoteSession.Availability -eq "Available") {
        $LastBackupSuccess = invoke-command -session $RemoteSession -scriptblock {
            [ARRAY]$arrWBAdminBackupVer = wbadmin get versions
            [STRING] $strVersionID = $arrWBAdminBackupVer[-3] #The third line from the end contains the date of the last successful backup
            [STRING] $strVersionID = $strVersionID.Replace('Version identifier: ', '') #Remove label
            [ARRAY] $strVersionID = $strVersionID.Split('-') #Split off the time
            [STRING] $strVersionID = $strVersionID[0] #Split off the time
            $strVersionID #echo out just the last date
        }
    }
    Else{#The computer is not available, try to get result again.
        If($TryCount -lt 2){
            $TryCount += 1
            Start-Sleep -Seconds 10
            Return (Is-WBAdminBackupSuccess $RemoteToCompName $BackupStartDate $TryCount)
        }
    }
    
    Remove-PSSession -Session $RemoteSession

    If($LastBackupSuccess -eq $BackupStartDate.ToString("MM/dd/yyyy")){
        Return $true
    }
    Else {
        Return $false
    }
}

#////////////////////////////////////
#               Hyper-V
#///////////////////////////////////

#I cannot take credit for the below function.  However, I've seemed to lost the reference of where it came from.  Appologies to the author.
Function Change-VMState{
    param([STRING] $HyperVHost, [STRING] $HyperVGuest, [STRING] $NewState)

    $s = new-pssession -computername $HyperVHost
    
    If ($s.Availability -eq "Available") {
        invoke-command -session $s -scriptblock {
            param($Guest, $Command)

            $Vm = Get-WmiObject -Namespace root\virtualization  -Query "Select * From Msvm_ComputerSystem Where ElementName = '$Guest'"
            $VM_Service = get-wmiobject -namespace root\virtualization Msvm_VirtualSystemManagementService
            $ListofVMs = get-wmiobject -namespace root\virtualization Msvm_ComputerSystem -filter  “ElementName <> Name ”

            If ($VM.ElementName -ne $Guest){
                write-warning "Virtual machine does not exist!"
                break
            }

            #Setting the "state" for later actions	  
            If ($Command -eq 'start'){
                $state = "started"
                $stateId = 2
            }
            ElseIf ($Command -eq 'stop'){
                $state = "stoped"
                $stateId = 3
            }
            ElseIf ($Command -eq 'pause'){
                $state = "paused"
                $stateId = 32768
            }
            ElseIf ($Command -eq 'save'){
                $state = "saved"
                $stateId = 32769
            }
            ElseIf (($Command -eq 'shutdown') -and ($Guest -ne $NUL)){
                #$VM_Service = get-wmiobject -namespace root\virtualization Msvm_VirtualSystemManagementService
                $ListofVMs = get-wmiobject -namespace root\virtualization Msvm_ComputerSystem -filter  “ElementName <> Name ”
                $Vm = Get-WmiObject -Namespace root\virtualization  -Query "Select * From Msvm_ComputerSystem Where ElementName = '$Guest'"
                $ShutdownIC = Get-WmiObject -Namespace root\virtualization  -Query "Associators of {$Vm} Where AssocClass=Msvm_SystemDevice ResultClass=Msvm_ShutdownComponent"
                $ShutdownIC.InitiateShutdown("TRUE", "Automated shutdown via script")	
                $state = "shutdown initiated"
                #Do this break because of avoiding errors - not fine but working
                break
            }
            ElseIf (($Command -eq 'snapshot')  -and ($Guest -ne $NUL)){
                #This part took a long time to get it to work, since the state "snapshot" is not recognized, or
                #let's say - wrong recognized in some cases - the state-id "32771" did not work properly for us
                #$VM_Service = get-wmiobject -namespace root\virtualization Msvm_VirtualSystemManagementService
                #$ListofVMs = get-wmiobject -namespace root\virtualization Msvm_ComputerSystem -filter  “ElementName <> Name ”

                ForEach ($VM in [array] $ListOfVMs){
                    # $VM.ElementName 
                    # $VM.__PATH
                    IF ($VM.ElementName -eq $Guest){  
                        $VM_service.CreateVirtualSystemSnapShot($VM.__PATH)
                        $state = "snapshot initiated"
                        $stateId = (32771)
                    }
                }
                $Guest + " " + $state
                break
            }
            #Check if specified command is valid
            else{   
                write-warning "Unknown command."
                "Syntax: .\hvcmd.ps1 [command] [machinename]" 
                "Commands: start stop pause save stop shutdown snapshot"
                "Machinename: Name of the Virtual instance (for example: server01)"
                break 
            }


            #Get a handle to the VM object
            $Core = get-wmiobject -namespace root\virtualization -class Msvm_Computersystem -filter "ElementName = '$Guest'" 

            #Set the state
            $status = $Core.RequestStateChange($stateId) 

            #Actual state
            If ($status.ReturnValue -ne 32775){
                $Guest + " " + $state
            }
            ElseIf ($status.ReturnValue -eq 32775){    
                "Nothing changed - allready in this state."  
            }

        } -ArgumentList $HyperVGuest, $NewState
        
        Remove-PSSession -Session $s
    }
}

Function Run-HyperVBackup{
    param([OBJECT] $BackupJob)

    $BackupReturn = New-Object Object
    $BackupReturn | add-member NoteProperty Status 'SUCCESS'
    $BackupReturn | add-member NoteProperty DestSubFolder ''
    $BackupReturn | add-member NoteProperty DestFullPath ''
    $BackupReturn | add-member NoteProperty PreviousBackupSize (New-Object Object)
    $BackupReturn | add-member NoteProperty BackupSize (New-Object Object)
    $BackupReturn | add-member NoteProperty CommandOutput (New-Object System.Collections.ArrayList)
    $BackupReturn | add-member NoteProperty ErrorOutput (New-Object System.Collections.ArrayList)

    #Stop VM Guest
    Change-VMState $BackupJob.HyperV_Host $BackupJob.TargetServer 'Stop' | Out-Null
    
    #Sleep for 5 Minutes to wait for shutdown
    Start-Sleep -s 300
    
    $BackupReturn.DestSubFolder = "HyperV\$($BackupJob.TargetServer)"
    $BackupReturn.DestFullPath = "$($BackupJob.DestFileShare)\$($BackupReturn.DestSubFolder)"

    #Take the last version and keep it in case the backup fails
    If($BackupJob.UsePriorBackupIfFailure -eq $true){
        If(test-path $($BackupReturn.DestFullPath)){
            Rename-Item "$($BackupReturn.DestFullPath)" "$($BackupReturn.DestFullPath).History"
            $BackupReturn.PreviousBackupSize = (Get-DiskUsed "$($BackupReturn.DestFullPath).History")
        }
    }
    Else{
        Remove-Item -path "$($BackupReturn.DestFullPath)" -Force -Recurse
    }
    
    #Copy Files to Backup Share
    Foreach($VHDSourceFolder IN ($BackupJob.HyperV_VHDFolder).Split('|')){

        $VHDDestFolder = "$($BackupReturn.DestFullPath)\$($VHDSourceFolder.Replace(':', ''))"

        #Create the folder
        New-Item $VHDDestFolder -type directory -force | Out-Null

        #Copy the file
        $BackupReturn.CommandOutput = Robocopy "$VHDSourceFolder" "$VHDDestFolder" /E /NP

        #Set the backup return status
        If((Is-RobocopySuccess $BackupReturn.CommandOutput) -eq $false){
            $BackupReturn.Status = 'FAILURE'
        }
    }
    
    Change-VMState $BackupJob.HyperV_Host $BackupJob.TargetServer 'Start' | Out-Null
    
    If($BackupReturn.Status = 'SUCCESS'){
        If($BackupJob.UsePriorBackupIfFailure -eq $true){
            If(test-path "$($BackupReturn.DestFullPath).History"){
                Remove-Item -path "$($BackupReturn.DestFullPath).History" -Force -Recurse
            }
        }
    }
    Else{
        If($BackupJob.UsePriorBackupIfFailure -eq $true){
            If(test-path $BackupReturn.DestFullPath){
                Remove-Item -path "$($BackupReturn.DestFullPath)" -Force -Recurse
            }
            If(test-path "$($BackupReturn.DestFullPath).History"){
                Rename-Item "$($BackupReturn.DestFullPath).History" "$($BackupReturn.DestFullPath)"
            }
        }
    }

    $BackupReturn.BackupSize = (Get-DiskUsed "$($BackupReturn.DestFullPath)")


    Return $BackupReturn
}

Function Is-RobocopySuccess{
    param([ARRAY] $RobocopyOutput)

    [STRING] $LineNeg5Output = $RobocopyOutput[-5]
    [STRING] $LineNeg9Output = $RobocopyOutput[-9]
    
    If($LineNeg5Output.Length -ge 10){
        If($LineNeg5Output.Substring(0,10) -eq '   Files :'){
            $FilesLine = $LineNeg5Output
        }
    }
    
    If($LineNeg9Output.Length -ge 10 -AND $FileLine -eq $null){
        If($LineNeg9Output.Substring(0,10) -eq '   Files :'){
            $FilesLine = $LineNeg9Output
        }
    }
    
    If($FilesLine -ne $null){
        [STRING] $FailedFiles = $FilesLine.Substring(59,10)
        [INT] $CountFailedFiles = $FailedFiles.Trim()
    }

    If($CountFailedFiles -eq 0){
        Return $true
    }
    Else{
        Return $false
    }
}

#////////////////////////////////////
#      Generic Backup Management
#///////////////////////////////////
Function Run-Backup{
    param([OBJECT]$objBackupJob)

    [DateTime] $BackupStartTime = Get-Date

    $BackupClient = 'UNKNOWN'
    $BackupStatus = 'FAILURE'

    If($objBackupJob.HyperV_Host -eq ''){ #Windows Backup
        If (Is-WSBInstalled $objBackupJob.TargetServer){
            $BackupClient = 'Windows Server Backup'
            $BackupReturn = Run-WSBBackup $objBackupJob
        }
        Else{
            $BackupClient = 'WBAdmin'
            $BackupReturn = Run-WBAdminBackup $objBackupJob
        }
    }
    Else{ #HyperV Backup
        $BackupClient = 'HyperV'
        $BackupReturn = Run-HyperVBackup $objBackupJob
    }

    $DiskAvlPostBackup = Get-DiskAvl $objBackupJob.DestFileShare
    
    [DateTime] $BackupEndTime = Get-Date
    $BackupElapsedTime = ($BackupEndTime - $BackupStartTime)
    $DiskAvlPostBackup = Get-DiskAvl $objBackupJob.DestFileShare

    $objBackupStat = New-Object Object
    $objBackupStat | add-member NoteProperty ServerName "$($objBackupJob.TargetServer)"
    $objBackupStat | add-member NoteProperty BackupClientUsed "$($BackupClient)"
    $objBackupStat | add-member NoteProperty DestBackupRoot @($objBackupJob.DestFileShare)
    $objBackupStat | add-member NoteProperty DestSubfolder @($BackupReturn.DestSubFolder)
    $objBackupStat | add-member NoteProperty DestFullPath @($BackupReturn.DestFullPath)
    $objBackupStat | add-member NoteProperty Status "$($BackupReturn.Status)"
    $objBackupStat | add-member NoteProperty CommandOutput @($BackupReturn.CommandOutput)
    $objBackupStat | add-member NoteProperty ErrorOutput @($BackupReturn.ErrorOutput)
    $objBackupStat | add-member NoteProperty DiskUsed ($BackupReturn.BackupSize)
    $objBackupStat | add-member NoteProperty DiskAvlPostBackup ($DiskAvlPostBackup)
    $objBackupStat | add-member NoteProperty DiskUsedPreviousBackup ($BackupReturn.PreviousBackupSize)
    $objBackupStat | add-member NoteProperty StartTime ($BackupStartTime)
    $objBackupStat | add-member NoteProperty EndTime ($BackupEndTime)
    $objBackupStat | add-member NoteProperty ElapsedTime ($BackupElapsedTime)
    $objBackupStat | add-member NoteProperty BackupThroughPut "$(($objBackupStat.DiskUsed.Megabytes)/($BackupElapsedTime.TotalSeconds))"
    
    Return $objBackupStat
}

Function Copy-BackupTargetsToUSBDrive {
    param(  [STRING] $USBDriveFullFilePath,
            [OBJECT] $TargetExecution,
            [INT] $TryCount = 0)

    #Copy all backups to drive
    $USBDriveCopyToPath = "$USBDriveFullFilePath\$($TargetExecution.DestSubfolder)"
    [ARRAY]$RobocopyOutput = Robocopy "$($TargetExecution.DestFullPath)" "$USBDriveCopyToPath" /E /NP

    #Set the backup return status
    If((Is-RobocopySuccess $RobocopyOutput) -eq $false){
        If($TryCount -lt 2){
            $TryCount += 1
            Return (Copy-BackupTargetsToUSBDrive $USBDriveFullFilePath $TargetExecution $TryCount)
        }
    }
    Else{
        Return 'SUCCESS'
    }
}

#////////////////////////////////////
#               File Share Management
#///////////////////////////////////


Function Get-DiskAvl {
    param ([STRING] $path)
    
    $AnyDriveLetter = 65..90 | ForEach-Object { ([char]$_)+":" }
    
    If(Test-path $path){
        If($AnyDriveLetter -contains "$($path.Split(':')[0]):"){
            $RootPath = "$($path.Split(':')[0]):"
        }
        ElseIf (($path.length -ge 2) -AND ($path.Substring(0,2) -eq '\\')){
            $RootPath = "\\$($path.Split('\')[2])\$($path.Split('\')[3])"
        }
        Else{
            $RootPath = ''
        }
    }
    Else{
        $RootPath = ''
    }
    
    If($RootPath -ne ''){
        $fso = new-Object -com Scripting.FileSystemObject
        $drv = $fso.getdrive($RootPath)
        $AvailableSpace = $drv.AvailableSpace
    }
    Else{
        $AvailableSpace = 0
    }
    
    $objSpace = New-Object Object
    $objSpace | add-member NoteProperty Bytes "$AvailableSpace"
    $objSpace | add-member NoteProperty Kilobytes "$($AvailableSpace/1024)"
    $objSpace | add-member NoteProperty Megabytes "$($AvailableSpace/1048576)"
    $objSpace | add-member NoteProperty Gigabytes "$($AvailableSpace/1073741824)"
        
    Return $objSpace
}


Function Get-DiskUsed {
    param ([STRING] $path)
    
    $Drive = Get-ChildItem $path -Recurse | Measure-Object -Property length -sum
    
    $objSpace = New-Object Object
    $objSpace | add-member NoteProperty Bytes "$($Drive.Sum)"
    $objSpace | add-member NoteProperty Kilobytes "$(($Drive.Sum)/1024)"
    $objSpace | add-member NoteProperty Megabytes "$(($Drive.Sum)/1048576)"
    $objSpace | add-member NoteProperty Gigabytes "$(($Drive.Sum)/1073741824)"
        
    Return $objSpace
}

Function Get-DriveWithRootDirectory {
    param ([STRING] $RootDirectory)

    $AvailableVolumes = gwmi Win32_volume | where {$_.BootVolume -ne $True -and $_.SystemVolume -ne $True}

    #Search all mounted volumes
    foreach ($AvailableVolume in $AvailableVolumes | Where {$_.DriveLetter -ne $null}){
        $TestBackupLocation = "$($AvailableVolume.DriveLetter)\$RootDirectory"
        If (Test-Path -path $TestBackupLocation){
            Return $TestBackupLocation
            Break
        }
    }

    #Search all un-mounted volumes
    foreach ($AvailableVolume in $AvailableVolumes | Where {$_.DriveLetter -eq $null}){
        
        $NewDriveLetter = Mount-Drive $AvailableVolume.DeviceID

        $TestBackupLocation = "$NewDriveLetter\$RootDirectory"
        If (Test-Path -path $TestBackupLocation){
            Return $TestBackupLocation
            Break
        }

        #Drive did not contain backup, dismount
        Dismount-Drive $NewDriveLetter
    }
}

Function Dismount-Drive{
    param([STRING] $DriveLetter)
    
    mountvol $DriveLetter /D

}

Function Mount-Drive{
    param([STRING] $VolumeDeviceID)
    
    $AnyDriveLetter = 65..90 | ForEach-Object { ([char]$_)+":" }
    $FreeDriveLetters = $AnyDriveLetter | Where-Object {(New-Object System.IO.DriveInfo($_)).DriveType -eq 'NoRootDirectory'}
    mountvol $FreeDriveLetters[0] $VolumeDeviceID
    
    Start-Sleep -seconds 5
    
    Return $FreeDriveLetters[0]
}

Function New-FileShare{
    param([STRING] $ShareName,
            [STRING] $ServerLocalPath)
    
    net share ""$ShareName""=""$ServerLocalPath"" "/GRANT:Everyone,FULL" | out-null
}

Function Delete-FileShare{
    param([STRING] $ShareName)
    
    Net share ""$ShareName"" /Delete /y
}

Function New-Email {
param ([STRING] $SMTPServer, [STRING] $emailFrom, [STRING] $emailTo, [STRING] $subject, [STRING] $body, [STRING] $attFile = '')

    $msg = new-object Net.Mail.MailMessage
    If ($attFile -ne "") { $att = new-object Net.Mail.Attachment($attFile) }
    $smtp = new-object Net.Mail.SmtpClient($SMTPServer)
    
    $msg.From = $emailFrom
    $msg.To.Add($emailTo)
    $msg.Subject = $subject
    $msg.Body = $body
    If ($attFile -ne "") { $msg.Attachments.Add($att) }
    
    $smtp.Send($msg)
    $msg.Attachments.Dispose()
}


#\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
#Create global objects
#/////////////////////////////

$gScriptPath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$gScriptName = $MyInvocation.MyCommand.Name

$gLogPath = $gScriptPath + '\Logs\'
New-Item $gLogPath -type directory -force | Out-Null



$gConfigFile = $gScriptPath + '\' + $gScriptName + '.xml'

[System.Xml.XmlDocument] $gXMLConfig = new-object System.Xml.XmlDocument
$gXMLConfig.load($gConfigFile)

#Load Globals
$gGlobals = $gXMLConfig.WindowsBackup.Globals

$BackupExecution = New-Object Object
$BackupExecution | add-member NoteProperty Schedules @()

$Schedules = $gXMLConfig.WindowsBackup.Schedules.Schedule | Where {$_.Enabled -ne 'NO'}
Foreach ($Schedule IN $Schedules){
    
    #Get Time Values for Backup Scheduling
    $SystemTime = (Get-Date).ToString("HHmm")
    $DayOfWeek = ((((Get-Date).DayOfWeek).ToString()).Substring(0,3)).ToUpper()
    $BackupWindow = ($Schedule.BackupWindow).Split("-")
    
    #Run backup if it is within the backup window
    If(($Schedule.AllowableDays).Contains($DayOfWeek) -AND ($SystemTime -gt $BackupWindow[0]) -AND ($SystemTime -lt $BackupWindow[1])){
        
        $ScheduleExecution = New-Object Object
        $ScheduleExecution | add-member NoteProperty ScheduleName ($Schedule.BackupScheduleName)
        $ScheduleExecution | add-member NoteProperty CopyToUSBStatus ''
        $ScheduleExecution | add-member NoteProperty CopyToUSBStartTime ''
        $ScheduleExecution | add-member NoteProperty CopyToUSBEndTime ''
        $ScheduleExecution | add-member NoteProperty CopyToUSBElapsedTime ''
        $ScheduleExecution | add-member NoteProperty CopyToUSBTotalDiskUsed (New-Object Object)
        $ScheduleExecution | add-member NoteProperty CopyToUSBTotalDiskAvl (New-Object Object)
        $ScheduleExecution | add-member NoteProperty CopyToUSBThroughPut ''
        $ScheduleExecution | add-member NoteProperty EntireErrorStatus ''
        $ScheduleExecution | add-member NoteProperty EmailSummary ''
        $ScheduleExecution | add-member NoteProperty LogFilePath ''
        $ScheduleExecution | add-member NoteProperty TargetExecutions @()
        
        $MarkerFile = $Schedule.BackupDestFileShare + '\' + (get-date).ToString("yyyyMMdd") + '.marker'
        
        If((test-path $MarkerFile) -eq $false){
            
            #Remove all other markers
            Remove-Item -path "$($Schedule.BackupDestFileShare)\*.marker"
            'Marker' | Out-File $MarkerFile
        
            $ScheduleExecution.LogFilePath = $gLogPath + $gScriptName + '_' + (get-date).ToString("yyyyMMddHHmmss") + ".txt"
            start-transcript -Path $ScheduleExecution.LogFilePath
            
            
            $Targets = $Schedule.Targets.Target | Where {$_.Enabled -ne 'NO'}
            Foreach($Target IN $Targets){

                #Clear all errors at this point            
                $error.clear()

                $objWB = New-Object Object
                $objWB | add-member NoteProperty TargetServer "$($Target.BackupTargetName)"
                $objWB | add-member NoteProperty UserName "$($Schedule.BackupUserName)"
                $objWB | add-member NoteProperty Password "$($Schedule.BackupPassword)"
                $objWB | add-member NoteProperty DestFileShare "$($Schedule.BackupDestFileShare)"
                $objWB | add-member NoteProperty BareMetalRecovery "$(
                                                                        If($Target.BareMetalRecovery -eq 'NO'){'NO'}
                                                                        Else{'YES'}
                                                                    )"
                $objWB | add-member NoteProperty VSSOptions "$(
                                                                If($Target.VSSOptions -eq 'COPY'){'COPY'}
                                                                Else{'FULL'}
                                                            )"
                $objWB | add-member NoteProperty AdditionalVolumes "$($Target.AdditionalVolumes)"
                $objWB | add-member NoteProperty HyperV_Host "$($Target.HyperV_Host)"
                $objWB | add-member NoteProperty HyperV_VHDFolder "$($Target.HyperV_VHDFolder)"
                $objWB | add-member NoteProperty MaxRetryCount "$(
                                                                    If($Target.MaxRetryCount -ne $null){$Target.MaxRetryCount}
                                                                    ElseIf ($Schedule.MaxRetryCount -ne $null){$Schedule.MaxRetryCount}
                                                                    Else {'1'} 
                                                                )"
                $objWB | add-member NoteProperty UsePriorBackupIfFailure ($(If($target.UsePriorBackupIfFailure -eq 'FALSE'){$false}
                                                                            Else{$true}))
                
                If($objWB.UsePriorBackupIfFailure -eq $false){
                    Write-Host "`r`n The option UsePriorBackupIfFailure has been set to false for $($objWB.TargetServer).  If an error occurs no backup of the server will be available."
                }
                
                #Process Max retries
                [INT]$BackupTry = 1
                Do {
                    Write-Host "`r`n Server: $($objWB.TargetServer) Try: $($BackupTry)/$($objWB.MaxRetryCount)"
                    Write-Host "`r`n"
                    $BackupResult = Run-Backup $objWB
                    $BackupResult | Add-Member NoteProperty BackupTryCount $BackupTry
                    
                    If($BackupResult.Status -eq 'SUCCESS'){break}
                    If($BackupTry -ge [INT]$objWB.MaxRetryCount){break}
                    If($BackupTry -ge [INT]10){break}
                    
                    $BackupTry += 1
                }
                While(1 -eq 1)

                #If the backup result is success, prior retry errored out.  Therefore, clear out errors.
                If($BackupResult.Status -eq 'SUCCESS'){
                    $error.clear()
                }
                Else{
                    #There was an error, add it to the error output
                    If($error){
                        $BackupResult.ErrorOutput += $error
                        $error.clear()
                    }
                }

                Write-Host "`r`nServer: $($BackupResult.ServerName)
                            `r`nBackup Client: $($BackupResult.BackupClientUsed)
                            `r`nStatus: $($BackupResult.Status)
                            `r`nBackup Try Count: $($BackupResult.BackupTryCount)/$($objWB.MaxRetryCount)
                            `r`nBare Metal Recovery: $($objWB.BareMetalRecovery)
                            `r`nAdditional Volumes: $($objWB.AdditionalVolumes)
                            `r`nBackup Size GBs: $("{0,0:n2}" -f [FLOAT]($BackupResult.DiskUsed.Gigabytes))
                            `r`nPrevious Backup Size GBs: $("{0,0:n2}" -f [FLOAT]($BackupResult.DiskUsedPreviousBackup.Gigabytes))
                            `r`nStart Time: $($BackupResult.StartTime)
                            `r`nEnd Time: $($BackupResult.EndTime)
                            `r`nElapsed Time: $($BackupResult.ElapsedTime)
                            `r`nBackup Through Put MBs/sec: $("{0,0:n2}" -f [FLOAT]($BackupResult.BackupThroughPut))
                            `r`n
                            `r`n-----------------------------------------------------------------------------
                            `r`n"

                $ScheduleExecution.TargetExecutions += $BackupResult
            
            }#Foreach target
            
            If($Schedule.USBDriveRootFolder -ne $null){
                $ScheduleExecution.CopyToUSBStartTime = Get-Date
            
                #Find the file path of the drive
                $USBDriveFullFilePath = Get-DriveWithRootDirectory $Schedule.USBDriveRootFolder
                Remove-Item -path "$USBDriveFullFilePath\*" -Force -Recurse
            
                #Copy the contents of the file share to the USB Drive
                $ScheduleExecution.CopyToUSBStatus = 'SUCCESS'
                Foreach($TargetExecution IN $ScheduleExecution.TargetExecutions){
                
                   $CopyToUSBStatus = Copy-BackupTargetsToUSBDrive $USBDriveFullFilePath $TargetExecution
                   If($CopyToUSBStatus -eq 'FAILURE'){
                        $ScheduleExecution.CopyToUSBStatus = 'FAILURE'
                   }
                }
            
                $ScheduleExecution.CopyToUSBEndTime = Get-Date
                $ScheduleExecution.CopyToUSBElapsedTime = $ScheduleExecution.CopyToUSBEndTime - $ScheduleExecution.CopyToUSBStartTime
                $ScheduleExecution.CopyToUSBTotalDiskUsed = Get-DiskUsed $USBDriveFullFilePath
                $ScheduleExecution.CopyToUSBTotalDiskAvl = Get-DiskAvl $USBDriveFullFilePath
                $ScheduleExecution.CopyToUSBThroughPut = "$(($ScheduleExecution.CopyToUSBTotalDiskUsed.Megabytes)/($ScheduleExecution.CopyToUSBElapsedTime.TotalSeconds))"
        

                #Determine entire error status of schedule execution
                $ScheduleExecution.EntireErrorStatus = 'SUCCESS'
        
                #Check USB Copy Status
                If($ScheduleExecution.CopyToUSBStatus -eq 'FAILURE'){
                    $ScheduleExecution.EntireErrorStatus = 'FAILURE'
                }
            
                #Check individual backup status
                If(@($ScheduleExecution.TargetExecutions | Where {$_.Status -eq 'FAILURE'}).Count -gt 0){
                    $ScheduleExecution.EntireErrorStatus = 'FAILURE'
                }
            
                #Check for scripting errors
                If ($error) {
                    $ScheduleExecution.EntireErrorStatus = 'FAILURE'
                }
            
                $ScheduleExecution.EmailSummary = "
                    Copy To USB Drive Status: $($ScheduleExecution.CopyToUSBStatus)
                    Copy To USB Drive Start Time: $($ScheduleExecution.CopyToUSBStartTime)
                    Copy To USB Drive End Time: $($ScheduleExecution.CopyToUSBEndTime)
                    Copy To USB Drive Elapsed Time: $($ScheduleExecution.CopyToUSBElapsedTime)
                    Copy To USB Drive Throughput: $("{0,0:n2}" -f [FLOAT]($ScheduleExecution.CopyToUSBThroughPut)) MB/s
                    USB Disk Available Post Backup: $("{0,0:n2}" -f [FLOAT]($ScheduleExecution.CopyToUSBTotalDiskAvl.Gigabytes)) GB
                "
            }#Should copy to USB Drive
            
            #Create Output for logging
            $ScheduleExecution.EmailSummary = "
                Destination Name: $($ScheduleExecution.ScheduleName)
                Completion Status: $($ScheduleExecution.EntireErrorStatus)
                Backup Start Time: $($ScheduleExecution.TargetExecutions[0].StartTime)
                Backup End Time: $($ScheduleExecution.TargetExecutions[-1].EndTime)
                Targets Backed Up: $($ScheduleExecution.TargetExecutions.Count)
                Targets With Errors: $(@($ScheduleExecution.TargetExecutions | Where {$_.Status -eq 'FAILURE'}).Count)
                Windows Backup Size: $("{0,0:n2}" -f [FLOAT]($ScheduleExecution.TargetExecutions | Select-Object -expandproperty DiskUsed | Measure-Object -Property Gigabytes -sum).Sum) GB
                File Share Disk Available Post Backup: $("{0,0:n2}" -f [FLOAT]($ScheduleExecution.TargetExecutions[-1].DiskAvlPostBackup.Gigabytes)) GB
                $($ScheduleExecution.EmailSummary)
               "
        
        
        
            Write-Host "`r`n$($ScheduleExecution.EmailSummary)"
            Write-Host "`r`n"
            Write-Host "`r`n                           Execution Detail                                  "
            Write-Host "`r`n-----------------------------------------------------------------------------"
            Write-Host "`r`n"
        
            Foreach($TargetExecution IN $ScheduleExecution.TargetExecutions){
                Write-Host "`r`n"
                Write-Host "`r`n-----------------------------------------------------------------------------"
                Write-Host "`r`n $($TargetExecution.ServerName)"
            
                Foreach($CommandOutputLine in $TargetExecution.CommandOutput){
                    Write-Host "`r`n $CommandOutputLine"
                }
            
                Write-Host "`r`n $($TargetExecution.ErrorOutput)"
            }
            
            #Add terminal error output
            If ($error) {
                Write-Host "`r`n"
                Write-Host "`r`n\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                Write-Host "`r`n//////////////////////ERRORS/////////////////////////"
                Write-Host "`r`n\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\"
                $error
                $error.clear()
            }
    
            stop-transcript

            If($Schedule.USBDriveRootFolder -ne $null){
                #Copy the log file to the drive
                Copy-Item $ScheduleExecution.LogFilePath $USBDriveFullFilePath

                #Dismount the backup drive
                Dismount-Drive "$($USBDriveFullFilePath.Split(':')[0]):"
            }

            $BackupExecution.Schedules += $ScheduleExecution
        
        }#Marker does not exist
    }#Within Window
}#Foreach Destination


#Email out results of executed schedules
Foreach($ExecutedSchedule In $BackupExecution.Schedules){
    If ($ExecutedSchedule.EntireErrorStatus -eq 'SUCCESS') {
         $EmailBody = "The backup completed successfully
         
                 $($ExecutedSchedule.EmailSummary)"
                 
        New-Email $gGlobals.EmailSMTP $gGlobals.EmailFrom $gGlobals.EmailTo "Windows Server Backup - Success" $EmailBody $ExecutedSchedule.LogFilePath
    }
    Else{
        $EmailBody = "The backup completed with errors.  Errors can be found at bottom of log.
        
                        $($ExecutedSchedule.EmailSummary)"
            
        New-Email $gGlobals.EmailSMTP $gGlobals.EmailFrom $gGlobals.EmailTo "Windows Server Backup - Failure" $EmailBody $ExecutedSchedule.LogFilePath
    }
}