#Requires -Modules @{ ModuleName="Az.Resources"; ModuleVersion="6.0.0" }
#Requires -Modules @{ ModuleName="Az.Accounts"; ModuleVersion="2.3.0" }
#Requires -Modules @{ ModuleName="Az.Automation"; ModuleVersion="1.7.0" }
#Requires -Modules @{ ModuleName="AzureAD"; ModuleVersion="2.0.2.140" }
#Requires -Modules @{ ModuleName="PnP.PowerShell"; ModuleVersion="1.12.0" }

Import-Module Az.Accounts -MinimumVersion 2.3.0
Import-Module Az.Automation -RequiredVersion 1.8.0
Import-Module PNP.PowerShell -MinimumVersion 1.12.0

. $PSScriptRoot\EIT-Teams-Provisioning-Helpers.ps1

Get-InstallationJSON

Create-AppRegistrations

Create-SharePointLandingSite
Apply-PnPTemplates

Create-AzureResources -Folder "ARMTemplates"

Configure-AutomationAccount -Folder "Runbooks"