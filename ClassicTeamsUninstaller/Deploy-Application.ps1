<#
.SYNOPSIS

PSApppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION

- The script is provided as a template to perform an install or uninstall of an application(s).
- The script either performs an "Install" deployment type or an "Uninstall" deployment type.
- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.

The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.

PSApppDeployToolkit is licensed under the GNU LGPLv3 License - (C) 2023 PSAppDeployToolkit Team (Sean Lillis, Dan Cunningham and Muhammad Mashwani).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the
Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but
WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
for more details. You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.

.PARAMETER DeploymentType

The type of deployment to perform. Default is: Install.

.PARAMETER DeployMode

Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.

.PARAMETER AllowRebootPassThru

Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.

.PARAMETER TerminalServerMode

Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Desktop Session Hosts/Citrix servers.

.PARAMETER DisableLogging

Disables logging to file for the script. Default is: $false.

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"

.EXAMPLE

powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"

.EXAMPLE

Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"

.INPUTS

None

You cannot pipe objects to this script.

.OUTPUTS

None

This script does not generate any output.

.NOTES

Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
- 69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
- 70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1

.LINK

https://psappdeploytoolkit.com
#>


[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall', 'Repair')]
    [String]$DeploymentType = 'Install',
    [Parameter(Mandatory = $false)]
    [ValidateSet('Interactive', 'Silent', 'NonInteractive')]
    [String]$DeployMode = 'Interactive',
    [Parameter(Mandatory = $false)]
    [switch]$AllowRebootPassThru = $false,
    [Parameter(Mandatory = $false)]
    [switch]$TerminalServerMode = $false,
    [Parameter(Mandatory = $false)]
    [switch]$DisableLogging = $false,
    #EIE Customizations
    [Parameter(Mandatory = $false)]
    [switch]$forceRemove = $false
)

Try {
    ## Set the script execution policy for this process
    Try {
        Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop'
    }
    Catch {
    }

    ##*===============================================
    ##* VARIABLE DECLARATION
    ##*===============================================
    ## Variables: Application
    [String]$appVendor = ''
    [String]$appName = 'Classic Teams Uninstaller'
    [String]$appVersion = ''
    [String]$appArch = ''
    [String]$appLang = 'EN'
    [String]$appRevision = '01'
    [String]$appScriptVersion = '1.0.0'
    [String]$appScriptDate = '02/09/2024'
    [String]$appScriptAuthor = 'Daniel Barton'
    ##*===============================================
    ## Variables: Install Titles (Only set here to override defaults set by the toolkit)
    [String]$installName = ''
    [String]$installTitle = ''

    ##* Do not modify section below
    #region DoNotModify

    ## Variables: Exit Code
    [Int32]$mainExitCode = 0

    ## Variables: Script
    [String]$deployAppScriptFriendlyName = 'Deploy Application'
    [Version]$deployAppScriptVersion = [Version]'3.9.3'
    [String]$deployAppScriptDate = '02/05/2023'
    [Hashtable]$deployAppScriptParameters = $PsBoundParameters

    ## Variables: Environment
    If (Test-Path -LiteralPath 'variable:HostInvocation') {
        $InvocationInfo = $HostInvocation
    }
    Else {
        $InvocationInfo = $MyInvocation
    }
    [String]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

    ## Dot source the required App Deploy Toolkit Functions
    Try {
        [String]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
        If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) {
            Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]."
        }
        If ($DisableLogging) {
            . $moduleAppDeployToolkitMain -DisableLogging
        }
        Else {
            . $moduleAppDeployToolkitMain
        }
    }
    Catch {
        If ($mainExitCode -eq 0) {
            [Int32]$mainExitCode = 60008
        }
        Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
        ## Exit the script, returning the exit code to SCCM
        If (Test-Path -LiteralPath 'variable:HostInvocation') {
            $script:ExitCode = $mainExitCode; Exit
        }
        Else {
            Exit $mainExitCode
        }
    }

    #endregion
    ##* Do not modify section above
    ##*===============================================
    ##* END VARIABLE DECLARATION
    ##*===============================================

    If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
        ##*===============================================
        ##* PRE-INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Pre-Installation'

        ## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
        Show-InstallationWelcome -CloseApps 'Teams=Classic Teams' -AllowDefer -DeferDeadline "03/21/2024" -PersistPrompt -CloseAppsCountdown 300

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Installation tasks here>
        # Ensure NewTeams is installed before uninstalling Classic Teams
    
        $MachineWide32 = Get-ChildItem -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue  | Get-ItemProperty | Where-Object { $_ -match 'Teams Machine-Wide Installer' }
        $MachineWide64 = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object { $_ -match 'Teams Machine-Wide Installer' }
    
        #In my environment a broken Teams Machine-Wide Installer had been pushed out via SCCM as a standalone install.  This caused a user-based installation without a "/ALLUSERS" install parameter.  In this case, the script will add in regkeys that MS has provided
        #to prevent the Teams Machine-Wide Installer from installing on any existing profiles
        #If your environment is different and you have isolated User Based Installs, you can update this block accordingly.

        if (($MachineWide32.InstallSource -like "C:\WINDOWS\ccmcache\*") -or ($MachineWide64.InstallSource -like "C:\WINDOWS\ccmcache\*")) {
            $userBased = $true
        }

            #if the deployment did not contain the -forceRemove parameter
            if ($forceRemove -eq $false) {

                Show-InstallationProgress -StatusMessage "Checking for New Teams..."
                $newTeams = Get-ChildItem -Path "C:\Program Files\WindowsApps\*" | Where-Object -Property "Name" -like "MSTeams*"

                if (!$newTeams) {

                    if ($deploymode -ne "Silent") {
                        $proceed = Show-InstallationPrompt -Message "New Teams is not detected on this system. Click ""Continue"" to proceed and uninstall Classic Teams leaving this device without any Teams client. Click ""Cancel"" to stop the uninstallation." -Title "WARNING: New Teams Not Detected" -ButtonLeftText "Cancel" -ButtonRightText "Continue" -PersistPrompt -Timeout 120 -ExitOnTimeout $true

                        if ($proceed -eq "Cancel") {
                            Show-InstallationPrompt -Message "Cancelling Uninstall - please contact your IT Support for help installing New Teams before continuing." -Icon "Information" -Title "Aborting Classic Teams Uninstall" -ButtonRightText "Ok"
                            exit-script -ExitCode 445
                        }

                    }
                    else {
                        Write-Log -Message "Error: Exiting Script without uninstalling Classic Teams.  User / Device does not have New Teams Installed.  Pass -forceRemove in your deployment parameters to override this and forcibly uninstall Classic Teams even if New Teams is not present."
                        exit-script -ExitCode 445
                    }

                }
            }
        

        ##*===============================================
        ##* INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Installation'

        ## Handle Zero-Config MSI Installations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) {
                $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ }
            }
        }
        ## <Perform Installation tasks here>
        # Function to uninstall Teams Classic

        
        Show-InstallationProgress -StatusMessage "Checking if Classic Teams is running..."
        try {
            $teamsProcess = Get-Process -Name "Teams" -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -like "*cannot find a process*") {
                Write-Log -Message "Classic Teams is not running"
            }
            else {
                write-host "Warning/Error: $_"
            }
        }

        if ($teamsProcess) {
            Show-InstallationProgress -StatusMessage "Closing Classic Teams..."
            Write-Log -Message "Teams is running, attempting to stop..."
            try {
                $teamsProcess | Stop-Process -Force -ErrorAction Stop
            }
            catch {
                Write-Log -Message "Warning/Error: $_"
            }
        }


        <# The following function is largely credited to FlorianSLZ with minor modifications to fit into PSADT
        https://github.com/FlorianSLZ/scloud/blob/main/Program%20-%20win32/Microsoft%20Teams%20(new)/install.ps1
        #>
        function Uninstall-TeamsClassic($TeamsPath) {
            try {
                $process = Execute-Process -Path "$TeamsPath\Update.exe" -Parameters "--uninstall /s" -PassThru -ErrorAction Stop -IgnoreExitCodes "*"

                if ($process.ExitCode -ne 0) {
                    Write-Log -Message "Uninstallation failed with exit code $($process.ExitCode)."
                }
            }
            catch {
                Write-Log -Message "ERROR/WARNING - $($_.Exception.Message)"
            }
        }

        # Get all Users
        Show-InstallationProgress -StatusMessage "Searching for Classic Teams installations, please wait..."
        $AllUsers = Get-ChildItem -Path "$($ENV:SystemDrive)\Users"

        # Process all Users
        foreach ($User in $AllUsers) {
            Write-Log -Message "Processing user: $($User.Name)"

            # Locate installation folder
            $localAppData = "$($ENV:SystemDrive)\Users\$($User.Name)\AppData\Local\Microsoft\Teams"
            $programData = "$($env:ProgramData)\$($User.Name)\Microsoft\Teams"

            if (Test-Path "$localAppData\Current\Teams.exe") {
                Write-Log -Message "Uninstalling Teams for user $($User.Name)"
                Show-InstallationProgress -StatusMessage "Found Classic Teams, please wait while it is uninstalled."
                Uninstall-TeamsClassic -TeamsPath $localAppData
            }
            elseif (Test-Path "$programData\Current\Teams.exe") {
                Write-Log -Message "Uninstall Teams for user $($User.Name)"
                Show-InstallationProgress -StatusMessage "Found Classic Teams, please wait while it is uninstalled."
                Uninstall-TeamsClassic -TeamsPath $programData
            }
            else {
                Show-InstallationProgress -StatusMessage "Still searching, please wait..."
                Write-Log -Message "Teams installation not found for user $($User.Name)"
            }
        }

        # Remove old Teams folders and icons
        Show-InstallationProgress -StatusMessage "Attempting to remove Classic Teams shortcuts and folders, please wait."
        Write-Log -Message "Attempting to remove Classic Teams shortcuts and folders from user profiles."

        $TeamsFolder_old = "$($ENV:SystemDrive)\Users\*\AppData\Local\Microsoft\Teams"
        $TeamsIcon_old = "$($ENV:SystemDrive)\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams*.lnk"

        Get-Item $TeamsFolder_old | Remove-Item -Force -Recurse
        Get-Item $TeamsIcon_old | Remove-Item -Force -Recurse

        #End FlorianSLZ credit

        if ($userBased -ne $true) {

            Show-InstallationProgress -StatusMessage "Searching for Teams Machine-Wide Installer..."
            Write-Log -Message "Attempting to find Teams Machine-wide Installer uninstall string."
        
            if ($MachineWide32) {

                $MsiPkg32Guid = "{39AF0813-FA7B-4860-ADBE-93B9B214B914}"
                Show-InstallationProgress -StatusMessage "Found 32-Bit Machine-Wide Installer, attempting to remove..."
                Write-Log -Message "Teams Machine Wide uninstall 32-Bit string found, attempting to silently uninstall."
                # Execute the uninstall command
                Execute-MSI -Action "Uninstall" -Path "$MsiPkg32Guid" -SkipMSIAlreadyInstalledCheck -IncludeUpdatesAndHotfixes -PassThru

            }
            else {

                Show-InstallationProgress -StatusMessage "No 32 bit Machine-Wide Installer Found..."
                Write-Host "32 Bit Teams Machine-Wide Installer not found"

            }

            if ($MachineWide64) {

                $MsiPkg64Guid = "{731F6BAA-A986-45A4-8936-7C3AAAAA760B}"
                Show-InstallationProgress -StatusMessage "Found 64-Bit Machine-Wide Installer, attempting to remove..."
                Write-Log -Message "Teams Machine Wide uninstall 64-Bit string found, attempting to silently uninstall."
                # Execute the uninstall command
                Execute-MSI -Action "Uninstall" -Path "$MsiPkg64Guid" -SkipMSIAlreadyInstalledCheck -IncludeUpdatesAndHotfixes -PassThru

            }
            else {
                Show-InstallationProgress -StatusMessage "No 64 bit Machine-Wide Installer Found..."
                Write-Host "64 Bit Teams Machine-Wide Installer not found"
            }
        }
        else {
            Show-InstallationProgress -StatusMessage "Removing Teams Machine-Wide Installer's functionality, please wait..."
            Write-Log -Message "A user based Teams Machine-Wide Installer was found, unable to install. Attempting to write registry keys to prevent it from functioning for all users"

            #Set registry key for all users to prevent Teams Machine Wide From Re-installing 
            [ScriptBlock]$HKCURegistrySettings = {
                #Registry key MS uses to tell the machine wide installer not to do anything https://learn.microsoft.com/en-us/microsoftteams/msi-deployment#uninstallation
                Set-RegistryKey -Key 'HKCU\Software\Microsoft\Office\Teams' -Name 'PreventInstallationFromMsi' -Value 1 -Type DWord -SID $UserProfile.SID
            }

            try {
                Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
            }
            catch {
                Write-Log -Message "Error: $_"
            }

            Write-Log -Message "Attempting to set system registry key to force Classic Teams to uninstall from all user profiles if detected on login."
            #Set registry key to tell the system not to install Classic Teams but to attempt to uninstall it if it exists
            try { 
                Set-RegistryKey -Key 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32' -Name 'TeamsMachineUninstallerLocalAppData' -Value "%LOCALAPPDATA%\Microsoft\Teams\Update.exe --uninstall --msiUninstall" -Type ExpandString -ErrorAction stop
            }
            catch {
                write-Log -Message "Error: $_"
            }
        }
        
        #>

        ##*===============================================
        ##* POST-INSTALLATION
        ##*===============================================
        [String]$installPhase = 'Post-Installation'

        ## <Perform Post-Installation tasks here>
        
    
        ## Display a message at the end of the install
        If (-not $useDefaultMsi) {
            Show-InstallationPrompt -Message 'Classic Teams Uninstaller has finished.' -ButtonRightText 'OK' -Icon Information -NoWait
        }
    }
    ElseIf ($deploymentType -ieq 'Uninstall') {
        ##*===============================================
        ##* PRE-UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Pre-Uninstallation'

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Uninstallation tasks here>


        ##*===============================================
        ##* UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Uninstallation'

        ## Handle Zero-Config MSI Uninstallations
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }

        ## <Perform Uninstallation tasks here>


        ##*===============================================
        ##* POST-UNINSTALLATION
        ##*===============================================
        [String]$installPhase = 'Post-Uninstallation'

        ## <Perform Post-Uninstallation tasks here>


    }
    ElseIf ($deploymentType -ieq 'Repair') {
        ##*===============================================
        ##* PRE-REPAIR
        ##*===============================================
        [String]$installPhase = 'Pre-Repair'

        ## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
        Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

        ## Show Progress Message (with the default message)
        Show-InstallationProgress

        ## <Perform Pre-Repair tasks here>

        ##*===============================================
        ##* REPAIR
        ##*===============================================
        [String]$installPhase = 'Repair'

        ## Handle Zero-Config MSI Repairs
        If ($useDefaultMsi) {
            [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) {
                $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
            }
            Execute-MSI @ExecuteDefaultMSISplat
        }
        ## <Perform Repair tasks here>

        ##*===============================================
        ##* POST-REPAIR
        ##*===============================================
        [String]$installPhase = 'Post-Repair'

        ## <Perform Post-Repair tasks here>


    }
    ##*===============================================
    ##* END SCRIPT BODY
    ##*===============================================

    ## Call the Exit-Script function to perform final cleanup operations
    Exit-Script -ExitCode $mainExitCode
}

Catch {
    [Int32]$mainExitCode = 60001
    [String]$mainErrorMessage = "$(Resolve-Error)"
    Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
    Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
    Exit-Script -ExitCode $mainExitCode
}

