#Global Variables
$global:installer = $null
[string]$global:certPath = $null
$global:certPassword = $null
$global:EITTeamsProvisioning_AppRegistration_ClientId = $null
$global:EITTeamsProvisioning_AppRegistration_CertThumbprint = $null
$global:EITTeamsProvisioning_API_AppRegistration_ClientId = $null
$global:siteCollectionRelativeUrl = $null
$global:rootSiteUrl = $null
$global:tenantId = $null

#Read InstallationJSON into object file
function Get-InstallationJSON{
    $global:installer = (Get-Content "$($PSScriptRoot)\installation.json" -Raw) | ConvertFrom-Json
    Write-Host -ForegroundColor Cyan "DEPLOYMENT STARTED!!"
    $installer
}

function Create-AppRegistrations{
    #Create Cert Folder
    $global:certPath = Create-Folder -FolderName "Certificates"

    Connect-AzAccount -Tenant $installer.PrimaryDomain
    $environment = Get-AzContext
    $global:tenantId = $environment.Tenant.Id

    $appReg = Get-AzADApplication -DisplayName "$($installer.ResourceGroup)_EITTeamsProvisioning"

    if($appReg){
        #Call funcation to set Teams Provisioning App Registration values
        Set-EITTeamsProvisioning-AR -appReg $appReg
    }
    else{
        #Call funcation to create Teams Provisioning App Registration
        Create-EITTeamsProvisioning-AR
    }

    $appRegAPI = Get-AzADApplication -DisplayName "$($installer.ResourceGroup)_EITTeamsProvisioning_API"
    if($appRegAPI){
        #Call funcation to set Teams Provisioning API App Registration values
        Set-EITTeamsProvisioning-API-AR -appReg $appReg
    }
    else{
        #Call funcation to create Teams Provisioning API App Registration
        Create-EITTeamsProvisioning-API-AR
    }
}

#Generate a folder in the root of the Script if it doesn't exist
function Create-Folder{
    Param
    (
        [Parameter(Mandatory = $true)][string] $FolderName
    )

    # Generate the folder
    $folder = "$($PSScriptRoot)\$($folderName)"
    if (!(Test-Path $folder)) {
        New-Item -Path $folder -ItemType "directory"
        Write-Host -ForegroundColor Green "Cert Folder created at: '$($folder)'"
    }
    
    return $folder
}

#Sets global variables for Teams Provisioning App Registration
function Set-EITTeamsProvisioning-AR{
    Param
    (
        [Parameter(Mandatory = $true)] $appReg
    )

    $appRegistrationName = "$($installer.ResourceGroup)_EITTeamsProvisioning"

    Write-Host -ForegroundColor Cyan "'$($appRegistrationName)' Already Exists..."
    #Set Teams Provisioning ClientId to use later
    $global:EITTeamsProvisioning_AppRegistration_ClientId = $appReg.AppId 
    
    #Set Teams Provisioning Cert Thumbprint to use later
    foreach ($key in $appReg.KeyCredentials){
        if ($key.DisplayName -eq "CN=$($appRegistrationName)"){
            $global:EITTeamsProvisioning_AppRegistration_CertThumbprint = [System.Convert]::ToBase64String($key.CustomKeyIdentifier)
        }
    }
}

#Create App Registration for Teams Provisioning
function Create-EITTeamsProvisioning-AR{
    $appRegistrationName = "$($installer.ResourceGroup)_EITTeamsProvisioning"

    #Generate Password for Cert
    $UnsecurePassword = New-RandomPassword -Length 20
    $global:certPassword = ConvertTo-SecureString $UnsecurePassword -AsPlainText -Force
    Write-Host -ForegroundColor Cyan "Certificate Password: $($UnsecurePassword)"
    Write-Host ""

    # Create the App Registration and assign SP and Graph API permissions
    $sharePointPermissions = "Sites.FullControl.All"
    $graphPermissions = "AccessReview.ReadWrite.All", "Group.Create", "Group.ReadWrite.All", "GroupMember.ReadWrite.All", "Notes.ReadWrite.All", "Team.Create", "User.Read.All"
    Write-Host -ForegroundColor Cyan "Creating App Registration $($appRegistrationName)..."
    $global:EITTeamsProvisioning_AppRegistration = Register-PnPAzureADApp -ApplicationName $appRegistrationName `
        -Tenant $installer.PrimaryDomain `
        -OutPath $certPath `
        -CertificatePassword $certPassword `
        -GraphApplicationPermissions $graphPermissions `
        -SharePointApplicationPermissions $sharePointPermissions `
        -Store CurrentUser `
        -Interactive
    
    # Update the app registration with additional Exchange permissions
    if ($EITTeamsProvisioning_AppRegistration -and $EITTeamsProvisioning_AppRegistration.'AzureAppId/ClientId') {
        Write-Host -ForegroundColor Cyan "Assigning Exchange permissions to App Registration $($appRegistration.'AzureAppId/ClientId')"
        $exchangePermissions = "Exchange.ManageAsApp"
    }

    $global:EITTeamsProvisioning_AppRegistration_ClientId = $EITTeamsProvisioning_AppRegistration.'AzureAppId/ClientId'
    $global:EITTeamsProvisioning_AppRegistration_CertThumbprint = $EITTeamsProvisioning_AppRegistration.'Certificate Thumbprint'
    
    Write-Host -ForegroundColor Green "$($appRegistrationName) Created!"
    Write-Host ""
}

#Generates a new Password
function New-RandomPassword {
#LINK - https://thesysadminchannel.com/generate-strong-random-passwords-using-powershell/
param(
        [Parameter(
            Position = 0,
            Mandatory = $false
        )]
        [ValidateRange(5,79)]
        [int]    $Length = 16,
    
        [switch] $ExcludeSpecialCharacters
    )
    
    $SpecialCharacters = @((33,35) + (36..38) + (42..44) + (60..64) + (91..94))
    try {
        if (-not $ExcludeSpecialCharacters) {
                $Password = -join ((48..57) + (65..90) + (97..122) + $SpecialCharacters | Get-Random -Count $Length | foreach {[char]$_})
            } else {
                $Password = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count $Length | foreach {[char]$_})
        }
    
    } catch {
        Write-Error $_.Exception.Message
    }

    return $Password
}

#Sets global variables for Teams Provisioning API App Registration
function Set-EITTeamsProvisioning-API-AR{
    Param
    (
        [Parameter(Mandatory = $true)] $appReg
    )
    $appRegistrationName = "$($installer.ResourceGroup)_EITTeamsProvisioning_API"

    Write-Host -ForegroundColor Cyan "'$($appRegistrationName)' Already Exists..."
    #Set Teams Provisioning ClientId to use later
    $global:EITTeamsProvisioning_API_AppRegistration_ClientId = $appReg.AppId
}

#Create App Registration for Teams Provisioning API
function Create-EITTeamsProvisioning-API-AR{
    $appRegistrationName = "$($installer.ResourceGroup)_EITTeamsProvisioning_API"

    # Create the App Registration and assign Graph API permissions
    $graphPermissions = "User.Read"
    Write-Host -ForegroundColor Cyan "Creating App Registration $($appRegistrationName)..."
    $global:EITTeamsProvisioning_API_AppRegistration = Register-PnPAzureADApp -ApplicationName $appRegistrationName `
        -Tenant $installer.PrimaryDomain `
        -GraphDelegatePermissions $graphPermissions `
        -Store CurrentUser `
        -Interactive

    #Set user_impersonation scope on provisioning API App Registration
    Set-AppRegistrationScope -ScopeName "user_impersonation" -AppId $EITTeamsProvisioning_API_AppRegistration.'AzureAppId/ClientId'

    $global:EITTeamsProvisioning_API_AppRegistration_ClientId = $EITTeamsProvisioning_API_AppRegistration.'AzureAppId/ClientId'

    Write-Host -ForegroundColor Green "$($appRegistrationName) Created!"
    Write-Host ""
}

#Set a scope on an app registration
function Set-AppRegistrationScope{
    Param
    (
        [Parameter(Mandatory = $true)][string] $ScopeName,
        [Parameter(Mandatory = $true)][string] $AppId
    )

    Write-Host "Setting Scope $ScopeName..."

    $adConnection = Connect-AzureAD -Domain $installer.PrimaryDomain

    $appCfg = Get-AzureADMSApplication -filter "AppId eq '$($AppId)'"

    $ScopeName = "user_impersonation"
    # Create access_as_user scope
    # Add all existing scopes first
    $scopes = New-Object System.Collections.Generic.List[Microsoft.Open.MsGraph.Model.PermissionScope]
    if ($appCfg.Api.Oauth2PermissionScopes -ne $null -and $appCfg.Api.Oauth2PermissionScopes.Count -gt 0) {
        $appCfg.Api.Oauth2PermissionScopes | foreach-object { $scopes.Add($_) }
    }
    $scope = CreateScope -value $ScopeName  `
        -userConsentDisplayName "" `
        -userConsentDescription "" `
        -adminConsentDisplayName "Envision IT Teams Provisioning API"  `
        -adminConsentDescription "Envision IT Teams Provisioning API to call the Logic App from the SPFx Site Request web part "
    $scopes.Add($scope)
    $appCfg.Api.Oauth2PermissionScopes = $scopes
    Set-AzureADMSApplication -ObjectId $appCfg.Id -Api $appCfg.Api
    Write-Host -ForegroundColor Green "Scope $ScopeName added."
}

#This function creates a new Azure AD scope (OAuth2Permission) with default and provided values
function CreateScope { 
Param
    (
        [string] $value,
        [string] $userConsentDisplayName,
        [string] $userConsentDescription,
        [string] $adminConsentDisplayName,
        [string] $adminConsentDescription
    )
    $scope = New-Object Microsoft.Open.MsGraph.Model.PermissionScope
    $scope.Id = New-Guid
    $scope.Value = $value
    $scope.UserConsentDisplayName = $userConsentDisplayName
    $scope.UserConsentDescription = $userConsentDescription
    $scope.AdminConsentDisplayName = $adminConsentDisplayName
    $scope.AdminConsentDescription = $adminConsentDescription
    $scope.IsEnabled = $true
    $scope.Type = "Admin"
    return $scope
}

#Create SharePoint Site Collection
function Create-SharePointLandingSite{

    #Set the Root and Relative URLs for the SharePoint site for later
    $global:siteCollectionRelativeUrl = $installer.SiteAddress -replace "https://.*sharepoint.com/", "/"
    $global:rootSiteUrl = $installer.SiteAddress -replace $siteCollectionRelativeUrl, ""

    #Connect to SharePoint Admin
    $adminSiteUrl = $rootSiteUrl -replace ".sharepoint.com", "-admin.sharepoint.com"
    Write-Host -ForegroundColor Yellow "Connecting to $adminSiteUrl - please connect with an account that has SharePoint admin rights"
    Connect-PnPOnline -Url $adminSiteUrl -Interactive

    Try{
        Write-Host "Checking if $($installer.SiteAddress) already exists..."

        $site = Get-PnPTenantSite -Url $installer.SiteAddress -ErrorAction SilentlyContinue
    }
    Catch{    
    }

    # Check the tenant recycle bin to see if the site is there
    if ($site -eq $null) {
        $tenantRecycleBin = Get-PnPTenantRecycleBinItem
        foreach ($recycledSite in $tenantRecycleBin) {
            if ($recycledSite.Url -eq $installer.SiteAddress) {
                # If it does it needs to be permanently deleted first
                $deleteSiteCollection = Read-Host "$($installer.SiteAddress) already exists in the tenant recycle bin. Permanently delete it (Y/N)?"
                if ($deleteSiteCollection.ToLower() -eq "y") {
                    Clear-PnPTenantRecycleBinItem -Url $installer.SiteAddress -Force
                }
            }
        }    
    }

    #Create SharePoint Site if it doesn't exist
    if ($site -eq $null) {
        Write-Host "$($installer.SiteAddress) doesn't exist, creating..."
        $siteTitle = Read-Host "Enter the title for the new site collection (Default: Landing)"
        if ($siteTitle -eq "") {
            $siteTitle = "Landing"
        }

        $site = New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $installer.SiteAddress -ErrorAction Stop
        Write-Host -ForegroundColor Green "$($installer.SiteAddress) has been created!"
        Write-Host ""
    }
    else{
        Write-Host -ForegroundColor Cyan "$($installer.SiteAddress) Already Exists..."
    }
    

    Disconnect-PnPOnline
}

#Apply PnP Templates to site collection and Upload to PnP Templates Library
function Apply-PnPTemplates{
    Write-Host -ForegroundColor Yellow "Connecting to $($installer.SiteAddress)"
    Connect-PnPOnline -Url $installer.SiteAddress -Interactive
    
    Write-Host "Applying Teams Provisioning Template to "$installer.SiteAddress
    Invoke-PnPSiteTemplate -Path "$PSScriptRoot\PnPTemplates\EUMSites.DeployTemplate.xml"
    Write-Host -ForegroundColor Green "Template Applied!"
    
    Write-Host "Uploading EUMSites.SiteMetadataList to PnP Templates library" 
    $siteMetaDataList = Add-PnPFile -Path "$PSScriptRoot\PnPTemplates\EUMSites.SiteMetadataList.xml" -Folder "PnPTemplates"
    $siteMetaDataOnlyList = Add-PnPFile -Path "$PSScriptRoot\PnPTemplates\EUMSites.SiteMetadataListOnly.xml" -Folder "PnPTemplates"
    Write-Host -ForegroundColor Green "Templates Uploaded!"
}

#Creates Azure Resources
function Create-AzureResources{
    Param(
        [Parameter(Mandatory = $true)] $Folder
    )
    
    Clear-AzContext -Force -ErrorAction SilentlyContinue
    Connect-AzAccount -Tenant $installer.PrimaryDomain
    Set-AzContext -Subscription $installer.SubscriptionId

    # Configure Logic Apps
    Write-Host -ForegroundColor Yellow "Starting Template Deployment for Azure Resources from Folder: $($Folder)"
    #Get the Folder and ARM Templates
    $templateFolder = "$($PSScriptRoot)\$($Folder)"
    $templates = Get-ChildItem "$($templateFolder)\" -Filter *.json

    $parameters = Set-AzureResourceParameters

    #Loop through all the ARM Templates in the Folder
    foreach ($template in $templates){
            #Call the ARMTemplate function to deploy the ARM Template
            DeployARMTemplate -TenantId $tenantId -SubscriptionId $installer.SubscriptionId -ResourceGroupName $installer.ResourceGroup -Location $installer.ResourceGroupRegion -TargetFolder $templateFolder `
                     -TemplateFilename "$($template.Name)" -TemplateParameters $parameters
    }

    Write-Host -ForegroundColor Green "Azure Resources Deployed!"
}

#Creates and returns a hash table to use for Azure ARM Template Deployment
function Set-AzureResourceParameters{
    $divisionListGUiD = Get-ListGUID -ListName "Divisions"
    $sitesListGUiD = Get-ListGUID -ListName "Sites"

    $parameters = @{
        "connections_azureautomation_name" = "azureautomation"
        "connections_office365_name"= "office365"
	    "connections_sharepointonline_name"= "sharepointonline"
        "connections_teams_name" = "teams"
		"workflows_CreateSite_SubmitForm_name" = "CreateSite-SubmitForm"
        "workflows_CreateSite_ProvisionSite_name" = "CreateSite-ProvisionSite"
        "workflows_ValidateBearerToken_name" = "ValidateBearerToken"
        "automationAccounts_eitdev_siteprovisioning_name" = "$($installer.AutomationAccountName)"
        "TenantId" = $tenantId
        "PrimaryDomain" = "$($installer.PrimaryDomain)"
        "ClientId" = $EITTeamsProvisioning_AppRegistration_ClientId
        "CertificateThumbprint" = $EITTeamsProvisioning_AppRegistration_CertThumbprint
        "TeamsProvisioningAPI_ClientId" = $EITTeamsProvisioning_API_AppRegistration_ClientId
		"SubscriptionId"= "$($installer.SubscriptionId)"
        "ResourceGroupName"= "$($installer.ResourceGroup)"
        "AutomationAccountName" = "$($installer.AutomationAccountName)"
        "Provision Site Logic App" = ""
        "RootURL" = $rootSiteUrl
        "SiteCollectionRelativeURL" = $siteCollectionRelativeUrl
        "Site Address" = $($installer.SiteAddress)
        "Divisions List GUID" = $divisionListGUiD
        "Sites List GUID" = $sitesListGUiD
        "Communication Method" = "$($installer.CommunicationMethod)"
        "Operator Email" = "$($installer.OperatorEmail)"
        "Teams Provisioning Approval Flow" = ""
        "SiteCollectionAdministrator" = ""
        "KeyVaultName" = ""
        "GuestInviterRoleGroup" = ""
        "EUMConfigSiteURL" = ""
        "B2BPolicyId" = ""
    }

    return $parameters
}

#Get the GUID for a SharePoint List
function Get-ListGUID{
    Param
    (
        [Parameter(Mandatory = $true)][string] $ListName
    )

    $List = Get-PnPList $ListName

    return $List.Id.Guid
}

#Deploys Azure Resources using ARMTemplate
function DeployARMTemplate {
    Param (
        [Parameter(Mandatory = $true)][string]  $TenantId,
        [Parameter(Mandatory = $true)][string]  $SubscriptionId,
        [Parameter(Mandatory = $true)][string]  $ResourceGroupName,
        [Parameter(Mandatory = $true)][string]  $Location,
        [Parameter(Mandatory = $true)][string]  $TargetFolder,
        [Parameter(Mandatory = $true)][String]  $TemplateFilename,
        [Parameter(Mandatory = $true)][Hashtable] $TemplateParameters
    )

    # Make sure $SubscriptionId is a valid Azure Subscription in this tenant.
    if ( Get-AzSubscription -SubscriptionId $SubscriptionId ) {
        # Subscription is valid.  We can proceed.
        Set-AzContext -SubscriptionId $SubscriptionId

        # Check if $ResourceGroupName exists in the selected Azure Subscription.
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorVariable notFound -ErrorAction SilentlyContinue
    
        if( $notFound ) {
            # If resource group does not exist, create it!
            Write-Host "Creating resource group '$ResourceGroupName'..."
            New-AzResourceGroup -Name $ResourceGroupName -Location $Location
        } else {
            Write-Host "Updating resource group '$ResourceGroupName'..."
        }

        # Deploy the full ARM Template in $ResourceGroupName.
        New-AzResourceGroupDeployment -ResourceGroupName $ResourceGroupName -TemplateFile "$TargetFolder\$TemplateFilename" -TemplateParameterObject $TemplateParameters

        # After deploying ARM Template...
        Write-Host "DONE: $($template.Name) ARM Deployment"
    }
    else {
        # Error handling: An invalid $SubscriptionId was provided
        Write-Host "Sorry, SubscriptionId '$SubscriptionId' was not found.  Please verify that the subscription exists and try again."
    }
}

#uploads Runbooks and Certificate to Automation Account
function Configure-AutomationAccount{
    Param(
        [Parameter(Mandatory = $true)] $Folder
    )

    Set-AzContext -SubscriptionId $installer.SubscriptionId

    #Upload Certificate to Automation Account
    Write-Host -ForegroundColor Yellow "Uploading Certificate to Azure Automation Account"
    $uploadedCertToAutomation = New-AzAutomationCertificate -AutomationAccountName $installer.AutomationAccountName -Name "$($installer.ResourceGroup)_EITTeamsProvisioning" -Path "$($certPath)\$($installer.ResourceGroup)_EITTeamsProvisioning.pfx" -Password $certPassword -ResourceGroupName $installer.ResourceGroup
    Write-Host -ForegroundColor Green "Certificate Uploaded!"

    $runbookFolder = "$($PSScriptRoot)\$($Folder)"
    $runbooks = Get-ChildItem "$($runbookFolder)\" -Filter *.ps1

    #Loop through all the ARM Templates in the Folder
    foreach ($runbook in $runbooks){
        $scriptNameWithoutExtension = $($runbook.Name).Replace('.ps1', '')
        Write-Host -ForegroundColor Yellow "Uploading runbook: $($runbook.Name)"
        Import-Runbooks -AutomationAccountName $installer.AutomationAccountName -ScriptName $scriptNameWithoutExtension -ResourceGroupName $installer.ResourceGroup -Type "PowerShell" -Path "$($runbookFolder)\$($runbook.Name)"
    }
    Write-Host -ForegroundColor Green "All Runbooks Uploaded!"
    
    Install-MicrosoftGraph-Modules -ResourceGroup $installer.ResourceGroup -AutomationAccount $installer.AutomationAccountName

    Install-AutomationAccount-Modules -ResourceGroup $installer.ResourceGroup -AutomationAccount $installer.AutomationAccountName

    Write-Host -ForegroundColor Green "DEPLOYMENT COMPLETE!!"
}

#Import Runbooks to Automation Account
function Import-Runbooks{
    Param
    (
        [Parameter(Mandatory = $true)][string] $AutomationAccountName,
        [Parameter(Mandatory = $true)][string] $ScriptName,
        [Parameter(Mandatory = $true)][string] $ResourceGroupName,
        [Parameter(Mandatory = $true)][string] $Type,
        [Parameter(Mandatory = $true)][string] $Path
    )

    $params = @{
        AutomationAccountName = $AutomationAccountName
        Name                  = $ScriptName
        ResourceGroupName     = $ResourceGroupName
        Type                  = $Type
        Path                  = $Path
    }

    Import-AzAutomationRunbook @params -Published -Force
}

#Install Microsoft Graph PowerShell module and dependencies
function Install-MicrosoftGraph-Modules{
    Param
    (
        [Parameter(Mandatory = $true)] $ResourceGroup,
        [Parameter(Mandatory = $true)] $AutomationAccount
    )

    Write-Verbose -Verbose -Message "Installing Microsoft Graph PowerShell module and dependencies"

    [System.Collections.Generic.List[Object]]$InstalledModules = @()
     
    #Get top level graph module
    $GraphModule = Find-Module Microsoft.Graph
    $DependencyList = $GraphModule | select -ExpandProperty Dependencies | ConvertTo-Json | ConvertFrom-Json
    $ModuleVersion = $GraphModule.Version
     
    #Since we know the authentication module is a dependency, let us get that one first
    $ModuleName = 'Microsoft.Graph.Authentication'
    $ContentLink = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$ModuleVersion"
    New-AzAutomationModule -ResourceGroupName $ResourceGroup -AutomationAccountName $AutomationAccount -Name $ModuleName -ContentLinkUri $ContentLink -ErrorAction Stop | Out-Null
    do {
        Start-Sleep 20
        $Status = Get-AzAutomationModule -ResourceGroupName $ResourceGroup -AutomationAccountName $AutomationAccount -Name $ModuleName | select -ExpandProperty ProvisioningState
    } until ($Status -in ('Failed','Succeeded'))
     
    if ($Status -eq 'Succeeded') {
        $InstalledModules.Add($ModuleName)
     
        foreach ($Dependency in $DependencyList) {
            $ModuleName = $Dependency.Name
            if ($ModuleName -notin $InstalledModules) {
                $ContentLink = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$ModuleVersion"
                New-AzAutomationModule -ResourceGroupName $ResourceGroup -AutomationAccountName $AutomationAccount -ContentLinkUri $ContentLink -Name $ModuleName -ErrorAction Stop | Out-Null
                sleep 3
            }
        }
     
        $LoopIndex = 0
        do {
            foreach ($Dependency in $DependencyList) {
                $ModuleName = $Dependency.Name
                if ($ModuleName -notin $InstalledModules) {
                    $Status = Get-AzAutomationModule -ResourceGroupName $ResourceGroup -AutomationAccountName $AutomationAccount -Name $ModuleName -ErrorAction SilentlyContinue | select -ExpandProperty ProvisioningState
                    sleep 3
                    if ($Status -in ('Failed','Succeeded')) {
                        if ($Status -eq 'Succeeded') {
                            $InstalledModules.Add($ModuleName)
                        }
     
                        [PSCustomObject]@{
                            Status           = $Status
                            ModuleName       = $ModuleName
                            ModulesInstalled = $InstalledModules.Count
                        }
                    }
                }
            }
            $LoopIndex++
        } until (($InstalledModules.Count -ge $GraphModule.Dependencies.count) -or ($LoopIndex -ge 10))
    }
}

#Install PowerShell modules and dependencies from AutomationModules.json
function Install-AutomationAccount-Modules{
    param(
        [Parameter(Mandatory = $true)] $ResourceGroup,
        [Parameter(Mandatory = $true)] $AutomationAccount
    )

    $automationModulesJSON = (Get-Content "$($PSScriptRoot)\automationModules.json" -Raw) | ConvertFrom-Json

    $automationModulesJSON.AutomationModules | ForEach-Object {
        $ModuleName = $_.ModuleName
        $ModuleVersion = $_.ModuleVersion
        Write-Host "Installing Module Name = $ModuleName $ModuleVersion ..."
        New-AzAutomationModule -AutomationAccountName $AutomationAccount -ResourceGroupName $ResourceGroup -Name $ModuleName -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/$ModuleName/$ModuleVersion"
        
        $Retries = 0;
        # Sleep for a few seconds to allow the module to become active (ordinarily takes a few seconds)
        Start-Sleep -s 30
        $ModuleReady = (Get-AzAutomationModule -AutomationAccountName $AutomationAccount -Name $ModuleName -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue).ProvisioningState

        # If new module is not ready, retry.
        While ($ModuleReady -ne "Succeeded" -and $Retries -le 6) {
            $Retries++;
            Write-Host "$ModuleName - Retry $Retries..."
            Start-Sleep -s 20
            $ModuleReady = (Get-AzAutomationModule -AutomationAccountName $AutomationAccount -Name $ModuleName -ResourceGroupName $ResourceGroup -ErrorAction SilentlyContinue).ProvisioningState
        }
        Write-Host ""
    }
}

