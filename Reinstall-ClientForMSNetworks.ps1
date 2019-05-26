<#
.SYNOPSIS
Reinstalls Client for Microsoft Networks.
Tested on Windows 7 SP1 with PowerShell 2.0

.Notes
Created by Willy Moselhy
Version 1.0 - May 26, 2019

Reinstalls Cleint for Microsoft Networks.
Tested on Windows 7 SP1 with PowerShell 2.0

The sample scripts provided here are not supported under any Microsoft standard support program or service. 
All scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties 
including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of 
the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, 
business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.DESCRIPTION
This script resintalles "Client for Microsoft Networks using the "NetCfg.exe" commands.
Script is designed for Windows 7.

By placing the script in the startup scripts, it will do the following on each restart,
1. Uninstall "Client for Microsoft Networks" and force reboot.
2. Install "Cleint for Microsoft Networks" and create a blocker file under the c:\

on all subsequent restarts until removed from startup it will check for the blocker file and terminate.

By default, logs are saved under C:\Windows\Temp


.EXAMPLE
.\Reinstall-ClientForMSNetworks.ps1

No parameters are needed for basic functionality. See complete help to make changes.

.LINK


#>
Param(
    [Parameter(Mandatory = $false)]
    # Path to store log file. Will create folder if it does not exist.
    [string] $LogFolder = "C:\Windows\Temp",

    [Parameter(Mandatory = $false)]
    # Show log on screen.
    [switch] $HostMode ,

    [Parameter(Mandatory = $false)]
    # Path to blocker file. If this file exists script will terminate.
    [string] $BlockerFilePath = "C:\Reinstall-ClientForMSNetworks.Block" 
)

    
#region: Script Configuration
    $ErrorActionPreference = "Stop"
    $ErrorThrown = $null
    $ScriptBeginTimeStamp = Get-Date
    $TimeStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $ErrorActionPreference = "Stop"
    $LogLevel = 0
    $VerMa="1"
    $VerMi="00"
    if($LogFolder){
        $LogFolderObject = New-Item -Path $LogFolder -ItemType Directory -Force
        $LogPath = "$LogFolderObject\ManageVMReplicationOnAVHDX_$TimeStamp.log"
    }
    if($LogPath){ # Logs will be saved to disk at the specified location
        $ScriptMode=$true
    }
    else{ # Logs will not be saved to disk
        $ScriptMode = $false
    }
#endregion: Script Configuration

#region: Logging Functions 
    #This writes the actual output - used by other functions
    function WriteLine ([string]$line,[string]$ForegroundColor, [switch]$NoNewLine){
        if($Script:ScriptMode){
            if($NoNewLine) {
                $Script:Trace += "$line"
            }
            else {
                $Script:Trace += "$line`r`n"
            }
            Set-Content -Path $script:LogPath -Value $Script:Trace
        }
        if($Script:HostMode){
            $Params = @{
                NoNewLine       = $NoNewLine -eq $true
                ForegroundColor = if($ForegroundColor) {$ForegroundColor} else {"White"}
            }
            Write-Host $line @Params
        }
    }
    
    #This handles informational logs
    function WriteInfo([string]$message,[switch]$WaitForResult,[string[]]$AdditionalStringArray,[string]$AdditionalMultilineString){
        if($WaitForResult){
            WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)$message" -NoNewline
        }
        else{
            WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)$message"  
        }
        if($AdditionalStringArray){
                foreach ($String in $AdditionalStringArray){
                    WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String"     
                }
       
        }
        if($AdditionalMultilineString){
            foreach ($String in ($AdditionalMultilineString -split "`r`n" | Where-Object {$_ -ne ""})){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String"     
            }
       
        }
    }

    #This writes results - should be used after -WaitFor Result in WriteInfo
    function WriteResult([string]$message,[switch]$Pass,[switch]$Success){
        if($Pass){
            WriteLine " - Pass" -ForegroundColor Cyan
            if($message){
                WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)`t$message" -ForegroundColor Cyan
            }
        }
        if($Success){
            WriteLine " - Success" -ForegroundColor Green
            if($message){
                WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)`t$message" -ForegroundColor Green
            }
        } 
    }

    #This write highlighted info
    function WriteInfoHighlighted([string]$message,[string[]]$AdditionalStringArray,[string]$AdditionalMultilineString){ 
        WriteLine "[$(Get-Date -Format hh:mm:ss)] INFO:    $("`t" * $script:LogLevel)$message"  -ForegroundColor Cyan
        if($AdditionalStringArray){
            foreach ($String in $AdditionalStringArray){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Cyan
            }
        }
        if($AdditionalMultilineString){
            foreach ($String in ($AdditionalMultilineString -split "`r`n" | Where-Object {$_ -ne ""})){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Cyan
            }
        }
    }

    #This write warning logs
    function WriteWarning([string]$message,[string[]]$AdditionalStringArray,[string]$AdditionalMultilineString){ 
        WriteLine "[$(Get-Date -Format hh:mm:ss)] WARNING: $("`t" * $script:LogLevel)$message"  -ForegroundColor Yellow
        if($AdditionalStringArray){
            foreach ($String in $AdditionalStringArray){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Yellow
            }
        }
        if($AdditionalMultilineString){
            foreach ($String in ($AdditionalMultilineString -split "`r`n" | Where-Object {$_ -ne ""})){
                WriteLine "[$(Get-Date -Format hh:mm:ss)]          $("`t" * $script:LogLevel)`t$String" -ForegroundColor Yellow
            }
        }
    }

    #This logs errors
    function WriteError([string]$message){
        WriteLine ""
        WriteLine "[$(Get-Date -Format hh:mm:ss)] ERROR:   $("`t`t" * $script:LogLevel)$message" -ForegroundColor Red
        
    }

    #This logs errors and terminated script
    function WriteErrorAndExit($message){
        WriteLine "[$(Get-Date -Format hh:mm:ss)] ERROR:   $("`t" * $script:LogLevel)$message"  -ForegroundColor Red
        Write-Host "Press any key to continue ..."
        $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
        $HOST.UI.RawUI.Flushinputbuffer()
        Throw "Terminating Error"
    }

#endregion: Logging Functions

#region: Script Functions
function Check-MSClientInstalled {
    $TempFilePath  = "C:\Windows\Temp\NETCFGOutput.tmp"
    $NetCFGProcess = Start-Process -FilePath NetCFG.exe -ArgumentList "-s n" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $TempFilePath
    $NetCFGOutput  = Get-Content -Path $TempFilePath | Where-Object {$_ -like "ms_msclient*"}
    if($NetCFGProcess.ExitCode -eq 0){
        if($NetCFGOutput){
            return $true
        }
        else{
            return $false
        }
    }
    else{
        throw "An error occured when checking for MSClient status. [NetCfg.exe -s n] $($NetCFGProcess.ExitCode)"
    }
}
function Uninstall-MSClient {
    $TempFilePath  = "C:\Windows\Temp\NETCFGOutput.tmp"
    $NetCFGProcess = Start-Process -FilePath NetCFG.exe -ArgumentList "-u MS_MSClient" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $TempFilePath
    if($NetCFGProcess.ExitCode -eq 0){
            return $true
    }
    else{
        throw "An error occured when uninstalling MSClient. [NetCfg.exe -u MS_MSClient] - $($NetCFGProcess.ExitCode)"
    }
}
function Install-MSClient {
    $TempFilePath  = "C:\Windows\Temp\NETCFGOutput.tmp"
    $NetCFGProcess = Start-Process -FilePath NetCFG.exe -ArgumentList "-c c -i MS_MSClient" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $TempFilePath
    if($NetCFGProcess.ExitCode -eq 0){
            return $true
    }
    else{
        throw "An error occured when installing MSClient. [NetCfg.exe -c c -i MS_MSClient] - $($NetCFGProcess.ExitCode)"
    }
}


#endregion: Script Functions


WriteInfo -message "Working on $Env:COMPUTERNAME - $(Get-Date -Format U)"
WriteInfo -message "Running as: $(whoami.exe)"
try{
    #region: Check for blocker file
    WriteInfo -message "ENTER: Check for blocker file"
    $LogLevel++
    
        if(Test-Path -Path $BlockerFilePath){
            throw "Blocker file found at '$BlockerFilePath'. Please stop running the script."
        }
        else{
            WriteInfo "Blocker file not found at '$BlockerFilePath'. Resuming script."
        }
    
    $LogLevel--
    WriteInfo -message "Exit:  Check for blocker file"
    #endregion: Check for blocker file
    
    #region: Check if msclient is installed
    WriteInfo -message "ENTER: Check if msclient is installed"
    $LogLevel++
        
        $MSClientInstalled = Check-MSClientInstalled
        WriteInfo -message "MSClient is installed: $MSClientInstalled"
    $LogLevel--
    WriteInfo -message "Exit:  Check if msclient is installed"
    #endregion: Check if msclient is installed
    
    if ($MSClientInstalled){ 
        #region: Uninstall MSClient and reboot
        WriteInfo -message "ENTER: Uninstall MSClient and reboot"
        $LogLevel++
        

            $MSClientUninstalled = Uninstall-MSClient
            if($MSClientInstalled){
                WriteInfo -message "MSClient uninstalled. Restarting!"
                Restart-Computer -Force
            }
        
        $LogLevel--
        WriteInfo -message "Exit:  Uninstall MSClient and reboot"
        #endregion: Uninstall MSClient and reboot
    }
    else {
        #region: Install MSClient and add blocker file
        WriteInfo -message "ENTER: Install MSClient and add blocker file"
        $LogLevel++
        
            $MSClientReinstalled = Install-MSClient
            if($MSClientReinstalled){
                WriteInfo -message "MSClient installed successfully."
                
                WriteInfo -message "Adding blocker file."
                Set-Content -Path $BlockerFilePath -Value "Created on $TimeStamp"
                
            }
        
        $LogLevel--
        WriteInfo -message "Exit:  Install MSClient and add blocker file"
        #endregion: Install MSClient and add blocker file
    }
}
#endregion: MAIN
catch{
    WriteError -message "An Error occured"
    WriteError -message $error[0].Exception.Message
    $ErrorThrown = $true
}
finally{
    $ScriptEndTimeStamp = Get-Date
    $LogLevel = 0
    WriteInfo -message "Script v$VerMa.$VerMi execution finished."
    Writeinfo -message "Duration: $(New-TimeSpan -Start $ScriptBeginTimeStamp -End $ScriptEndTimeStamp)"

    if($ErrorThrown) {
        Throw $error[0].Exception.Message
        exit(-1)
    }
    else{
        exit(0)
    }
}