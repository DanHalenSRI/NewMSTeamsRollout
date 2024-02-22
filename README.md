# NewMSTeamsRollout
PS App Deploy Toolkit Installer and Uninstaller for New / Classic Microsoft Teams

Credit to PSAppDeployToolkit for the PSADT files: https://github.com/psappdeploytoolkit/psappdeploytoolkit

Read More about using and deploying PSADT Scripts here: https://psappdeploytoolkit.com

## ClassicTeamsUninstaller

- When deploying via SCCM use "Deploy-Application.exe" as the "Installation Program"

- This is built as an interactive uninstaller, warning users if New Teams is not installed and allowing them the opportunity to cancel the program

- You can bypass that by using ""Deploy-Application.exe" -forceRemove" as the "Installation Program"

## NewTeamsInstaller

- This is designed to be a silent installation with no user prompts

- This will automatically try to download the teamsbootsrapper.exe from microsoft if the device has a network connection

- If no network connection it will revert to an offline installer

- Make sure you put the teamsbootstrapper.exe and the MSTeams-x64.msix apps in the Files directory before deploying so it can properly fallback if need be

- Auto detection of x86 / ARM architecture for offline fallback will be added shortly


Please feel free to provide feedback or pull requests 