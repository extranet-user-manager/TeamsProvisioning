# TeamsProvisioning
Teams Provisioning Installation V6.14

## Prerequistes
	PowerShell Modules
		Azure Az
			Must have Azure Az.Accounts and Azure Az.Resources
		Azure AD
			Version 2.0.2.140
		PnP.PowerShell
			Version 1.12.0
	Global Admin Account
	Owner Access to Azure Subscription
	PowerAutomate Premium
	
## Deployed Resources
	App Registrations
		<Resource Group Name>_EITTeamsProvisioning
		<Resource Group Name>_EITTeamsProvisioning_API
	Azure Resources
		1 - Azure Resource Group (Will be created if it doesn't exist)
		1 - Automation Account
		4 - Logic App API Connections
		3 - Logic Apps
		4 - Runbooks
	SharePoint Site Collection (Will be created if it doesn't exist)
	
## Sample Installaion.json
```
{
	"PrimaryDomain":"test.onmicrosoft.com",
	"SubscriptionId":"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
	"ResourceGroup":"Test-Resource-Group",
	"ResourceGroupRegion":"Canada Central",
	"SiteAddress":"https://test.sharepoint.com/sites/test-landing",
	"AutomationAccountName":"test-aa",
	"CommunicationMethod":"Email", (Can be set to 'Email or 'Teams')
	"OperatorEmail":"tester@test.onmicrosoft.com"
}
```

## Installation Steps
	Populate the installation.json file in the package
	Unblock both ps1 files by going to properties of each file and checking the checkbox if it exists
	Run EIT-Teams-Provisioning-Configuration.ps1
	Update the Submit_Form Logic App Parameter 'Provision Site Logic App' to the URL of the CreateSite-ProvisionSite Logic App
	Perform Manual Steps outlined in Teams Provisioning Deployment Guide