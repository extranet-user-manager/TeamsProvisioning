function CreateSite-Customizations {
    Param
    (
        [Parameter (Mandatory = $true)][int]$listItemID
    )

    Write-Verbose "CreateSite-Customizations Debug 1"

    $connLandingSite = Helper-Connect-PnPOnline -Url $SiteCollectionFullURL

    $pendingSiteCollection = Get-PnPListItem -Connection $connLandingSite -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='ID'/>
                    <Value Type='Integer'>$listItemID</Value>
                </Eq>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMSiteVisibility'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='EUMDivision'></FieldRef>
            <FieldRef Name='EUMCreateTeam'></FieldRef>
            <FieldRef Name='EUMSensitivityLabels'></FieldRef>
            <FieldRef Name='EUMLimitSharingByDomain'></FieldRef>
            <FieldRef Name='EUMLimitSharingDomains'></FieldRef>
            <FieldRef Name='EUMDefaultSharingLinkType'></FieldRef>
            <FieldRef Name='ExpectedCollaborationEndDate'></FieldRef>
            <FieldRef Name='EUMAdditionalOwners'></FieldRef>
            <FieldRef Name='EUMExternalSharing'></FieldRef>
            <FieldRef Name='Author'></FieldRef>
            <FieldRef Name='Requestor'></FieldRef>
        </ViewFields>
    </View>"

    if ($pendingSiteCollection.Count -eq 1) {
        [string]$eumSiteTemplate = $pendingSiteCollection["EUMSiteTemplate"].LookupValue
        [string]$siteURL = $pendingSiteCollection["EUMSiteURL"]
        [string]$portfolio = $pendingSiteCollection["EUMDivision"].LookupValue
        [string]$groupName = $pendingSiteCollection["Title"]
        [array]$eumAdditionalOwners = $pendingSiteCollection["EUMAdditionalOwners"]
        [string]$eumExternalSharing = $pendingSiteCollection["EUMExternalSharing"]
        [string]$requester = $pendingSiteCollection["Requestor"].Email

        $siteTemplate = Get-PnPListItem -Connection $connLandingSite -List "Site Templates" -Query "
                                                    <View>
                                                        <Query>
                                                            <Where>
                                                                <Eq>
                                                                    <FieldRef Name='Title'/>
                                                                    <Value Type='Text'>$eumSiteTemplate</Value>
                                                                </Eq>
                                                            </Where>
                                                        </Query>
                                                        <ViewFields>
                                                            <FieldRef Name='Title'></FieldRef>
                                                            <FieldRef Name='BaseClassicSiteTemplate'></FieldRef>
                                                            <FieldRef Name='BaseModernSiteType'></FieldRef>
                                                            <FieldRef Name='PnPSiteTemplate'></FieldRef>
                                                        </ViewFields>
                                                    </View>"
            
        $baseSiteTemplate = ""
        $baseSiteType = ""
        $pnpSiteTemplate = ""

        if ($siteTemplate.Count -eq 1) {
            $baseSiteTemplate = $siteTemplate["BaseClassicSiteTemplate"]
            $baseSiteType = $siteTemplate["BaseModernSiteType"]
            if ($siteTemplate["PnPSiteTemplate"] -ne $null) {
                $pnpSiteTemplate = "$pnpTemplatePath\$($siteTemplate["PnPSiteTemplate"].LookupValue)"
            }
        }
        
        [string]$eumDefaultSharingLinkType = $pendingSiteCollection["EUMDefaultSharingLinkType"]
        [string]$eumLimitSharingByDomain = $pendingSiteCollection["EUMLimitSharingByDomain"]
        [string]$eumLimitSharingDomains = $pendingSiteCollection["EUMLimitSharingDomains"]

        Write-Verbose -Verbose -Message "PnPTemplate: $($pnpSiteTemplate)"
        if (($pnpSiteTemplate -like "*Client-Template*.xml")) {
            Write-Verbose "Updating Client Site"
            Helper-Connect-PnPOnline -Url $siteURL

            $spFolder = Add-PnPFolder -Name "Quotes" -Folder "Shared Documents"
            $spFolder = Add-PnPFolder -Name "Signed Quotes" -Folder "Shared Documents/Quotes"
            $spFolder = Add-PnPFolder -Name "Invoices" -Folder "Shared Documents"

            $spFolder = Add-PnPFolder -Name "Business Development" -Folder "Private Documents"
            $spFolder = Add-PnPFolder -Name "Confidential" -Folder "Private Documents"
            $spFolder = Add-PnPFolder -Name "Quotes" -Folder "Private Documents"

            Remove-PnPContentTypeFromList -List "Shared Documents" -ContentType "Document"
            Remove-PnPContentTypeFromList -List "Private Documents" -ContentType "Document"

            Remove-QuickLaunchNode -Title "Conversations" -Url $siteURL
            Remove-QuickLaunchNode -Title "Pages" -Url $siteURL
            Remove-QuickLaunchNode -Title "Home" -Url $siteURL
            
            Add-QuickLaunchNode -Title "Teams" -SiteUrl $siteURL -First
            Add-QuickLaunchNode -Title "Home" -SiteUrl $siteURL -NodeUrl $siteURL  -First
        }
        elseif (($pnpSiteTemplate -like "*EUM-Committee-Template.xml")) {
            # Connect to the newly created site and get its group ID
            $connNewSite = Helper-Connect-PnPOnline -Url $siteURL
            $site = Get-PnPSite -Connection $connNewSite -Includes GroupId

            # Make sure we have access to the Azure Automation variables
            if (![string]::IsNullOrWhiteSpace($site.GroupId) -and $AzureAutomation) {
                # Add the M365 Group to the Visitors group and remove from the Members group
                $m365GroupClaim = "c:0o.c|federateddirectoryclaimprovider|$($site.GroupId)"
                $visitorGroupName = "$($groupName) Visitors"
                $memberGroupName = "$($groupName) Members"

                Write-Verbose -Verbose -Message "Adding $($m365GroupClaim) to $($visitorGroupName)"
                Add-PnPGroupMember -LoginName $m365GroupClaim -Group $visitorGroupName -Connection $connNewSite

                Write-Verbose -Verbose -Message "Removing $($m365GroupClaim) from $($memberGroupName)"
                Remove-PnPGroupMember -LoginName $m365GroupClaim -Group $memberGroupName -Connection $connNewSite                
                
                # Get and connect to the EUM Config site
                $eumConfigSite = Get-AutomationVariable -Name 'EUMConfigSiteURL'
                $connEUMConfigSite = Helper-Connect-PnPOnline -Url $eumConfigSite

                #Build the empty JSON group file
                $stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes("{`"Members`": `"`"}"))
                
                # Get the EUM page templates for the committee and meeting pages
                $committeeTemplate = Get-PnPListItem -List PublisherPageTemplates -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Group Member Committee Page</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $meetingTemplate = Get-PnPListItem -List PublisherPageTemplates -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Committee Meeting Page</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # Get the EUM values needed for webpart
                $eumAdminURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>AdminURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumPortalURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>PortalURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumAPIApplicationIDURI = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>EUM-API-Application-ID-URI</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # EUM page templates found
                if ($committeeTemplate -ne $null -and $committeeTemplate.Id -ne $null -and $meetingTemplate -ne $null -and $meetingTemplate.Id -ne $null) {
                    $folderName = $groupName.Replace(" ", "-").ToLower()
                    # Convert the Azure AD Group into an EUM group
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false; MemberPageTemplate = $committeeTemplate.Id; MemberPage = "Auto-generated URL"; MemberPageURL = "/members/$($folderName)" }
                    
                    # Create the members folder and meeting page
                    Add-PnPFolder -Name $folderName -Folder "PublisherPages/members" -Connection $connEUMConfigSite
                    $meetingStream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes(""))
                    Add-PnPFile -FileName "meeting.cshtml" -Folder "PublisherPages/members/$($folderName)" -Stream $meetingStream -Connection $connEUMConfigSite -Values @{Title = "Meeting"; ContentType = "EUM Base Page"; PublisherPageTemplate = $meetingTemplate.Id; GroupIDs = $site.GroupId } -Publish -PublishComment "Published by CreateSite-Customizations.ps1"
                }
                # No EUM page templates
                else {
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false }
                }
            }

            #Setup Documents Library of Committee site
            Write-Verbose -Verbose -Message "Creating Documents folder structure"
            Add-PnPFolder -Name "General" -Folder "Shared Documents" -Connection $connNewSite
            Add-PnPFolder -Name "Archived Documents" -Folder "Shared Documents/General" -Connection $connNewSite
            Add-PnPFolder -Name "Committee Documents" -Folder "Shared Documents/General" -Connection $connNewSite
            Add-PnPFolder -Name "Meeting Documents" -Folder "Shared Documents/General" -Connection $connNewSite

            # Apply the Home Page Canvas Content
            Write-Verbose -Verbose -Message "Applying EUM-Committee-Template home page"
            $canvasContent = Get-PnPFile -Url "$($SiteCollectionRelativeURL)/PnPTemplates/EUM-Committee-Template.html" -Connection $connLandingSite -AsString
            
            $eumAdminURLEncoded = $eumAdminURL -replace "https://", "https&#58;//"
            $eumPortalURLEncoded = $eumPortalURL -replace "https://", "https&#58;//"
            $eumAPIApplicationIDURIEncoded = $eumAPIApplicationIDURI -replace "api://", "api&#58;//"

            $canvasContent = $canvasContent -replace "~eumAdminURL~", $eumAdminURLEncoded
            $canvasContent = $canvasContent -replace "~eumPortalURL~", $eumPortalURLEncoded
            $canvasContent = $canvasContent -replace "~eumAPIApplicationIDURI~", $eumAPIApplicationIDURIEncoded

            $pageNameWithExtension = "Home.aspx"
            $page = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='String'>$($pageNameWithExtension)</Value></Eq></Where></Query></View>" -Connection $connNewSite
            $setPage = Set-PnPListItem -List SitePages -Identity $page.Id -Values @{"CanvasContent1" = $canvasContent } -UpdateType SystemUpdate -Connection $connNewSite
        }
        elseif (($pnpSiteTemplate -like "*EUM-Project-Template.xml")) {
            # Connect to the newly created site and get its group ID
            $connNewSite = Helper-Connect-PnPOnline -Url $siteURL
            $site = Get-PnPSite -Connection $connNewSite -Includes GroupId

            # Make sure we have access to the Azure Automation variables
            if (![string]::IsNullOrWhiteSpace($site.GroupId) -and $AzureAutomation) {
                # Get and connect to the EUM Config site
                $eumConfigSite = Get-AutomationVariable -Name 'EUMConfigSiteURL'
                $connEUMConfigSite = Helper-Connect-PnPOnline -Url $eumConfigSite

                #Build the empty JSON group file
                $stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes("{`"Members`": `"`"}"))
                
                # Get the EUM page templates for the pproject member page
                $projectTemplate = Get-PnPListItem -List PublisherPageTemplates -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Group Member Project Page</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # Get the EUM values needed for webpart
                $eumAdminURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>AdminURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumPortalURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>PortalURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumAPIApplicationIDURI = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>EUM-API-Application-ID-URI</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # EUM page templates found
                if ($projectTemplate -ne $null -and $projectTemplate.Id -ne $null) {
                    $folderName = $groupName.Replace(" ", "-").ToLower()
                    # Convert the Azure AD Group into an EUM group
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false; MemberPageTemplate = $projectTemplate.Id; MemberPage = "Auto-generated URL"; MemberPageURL = "/members/$($folderName)" }
                }
                # No EUM page templates
                else {
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false }
                }
            }

            #Setup Documents Library of Project site
            Write-Verbose -Verbose -Message "Creating Documents folder structure"
            Add-PnPFolder -Name "General" -Folder "Shared Documents" -Connection $connNewSite
            Add-PnPFolder -Name "External Groups" -Folder "Shared Documents/General" -Connection $connNewSite

            # Apply the Home Page Canvas Content
            Write-Verbose -Verbose -Message "Applying EUM-Project-Template home page"
            $canvasContent = Get-PnPFile -Url "$($SiteCollectionRelativeURL)/PnPTemplates/EUM-Project-Template.html" -Connection $connLandingSite -AsString
            
            $eumAdminURLEncoded = $eumAdminURL -replace "https://", "https&#58;//"
            $eumPortalURLEncoded = $eumPortalURL -replace "https://", "https&#58;//"
            $eumAPIApplicationIDURIEncoded = $eumAPIApplicationIDURI -replace "api://", "api&#58;//"
            
            $canvasContent = $canvasContent -replace "~eumAdminURL~", $eumAdminURLEncoded
            $canvasContent = $canvasContent -replace "~eumPortalURL~", $eumPortalURLEncoded
            $canvasContent = $canvasContent -replace "~eumAPIApplicationIDURI~", $eumAPIApplicationIDURIEncoded
            
            $pageNameWithExtension = "Home.aspx"
            $page = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='String'>$($pageNameWithExtension)</Value></Eq></Where></Query></View>" -Connection $connNewSite
            $setPage = Set-PnPListItem -List SitePages -Identity $page.Id -Values @{"CanvasContent1" = $canvasContent } -UpdateType SystemUpdate -Connection $connNewSite            
        }
        elseif (($pnpSiteTemplate -like "*EUM-Data-Room-Template.xml")) {
            # Connect to the newly created site and get its group ID
            $connNewSite = Helper-Connect-PnPOnline -Url $siteURL
            $site = Get-PnPSite -Connection $connNewSite -Includes GroupId,RootWeb

            # Make sure we have access to the Azure Automation variables
            if (![string]::IsNullOrWhiteSpace($site.GroupId) -and $AzureAutomation) {
                # Get and connect to the EUM Config site
                Write-Output -Verbose -Message "Get and connect to the EUM Config site"

                $eumConfigSite = Get-AutomationVariable -Name 'EUMConfigSiteURL'
                $connEUMConfigSite = Helper-Connect-PnPOnline -Url $eumConfigSite

                # Get the EUM page templates for the pproject member page
                $dataRoomTemplate = Get-PnPListItem -List PublisherPageTemplates -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Group Member Document Sharing</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # Get the EUM values needed for webpart
                $eumAdminURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>AdminURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumPortalURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>PortalURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumAPIApplicationIDURI = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>EUM-API-Application-ID-URI</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                #Copy Templates folder and contents to new site with metadata
                Add-PnPFolder -Name "Templates" -Folder "Shared Documents/General" -Connection $connNewSite
                Copy-PnPFile -SourceUrl "$($SiteCollectionRelativeURL)/PnPTemplates/Document Tracking.json" -TargetUrl "Shared Documents/General/Templates" -IgnoreVersionHistory -Connection $connNewSite -Force
                #Wait 10 seconds for copy to complete
                Start-Sleep -Seconds 10
                $DocumentJson = Get-PnPFile -Url "Shared Documents/General/Templates/Document Tracking.json" -AsListItem -Connection $connNewSite
                Set-PnPListItem -List "Documents" -Identity $DocumentJson.Id -Values @{"InternalDocumentType" = "Document Tracking JSON"} -Connection $connNewSite          
                
                Copy-PnPFile -SourceUrl "$($SiteCollectionRelativeURL)/PnPTemplates/Document Tracking.xlsx" -TargetUrl "Shared Documents/General/Templates" -IgnoreVersionHistory -Connection $connNewSite -Force
                #Wait 10 seconds for copy to complete            
                Start-Sleep -Seconds 10
                $DocumentExcel = Get-PnPFile -Url "Shared Documents/General/Templates/Document Tracking.xlsx" -AsListItem -Connection $connNewSite  
                Set-PnPListItem -List "Documents" -Identity $DocumentExcel.Id -Values @{"InternalDocumentType" = "Document Tracking Excel"} -Connection $connNewSite          
                
                # Create new permission level Contribute No Delete
                Add-PnPRoleDefinition -RoleName "Contribute No Delete" -Clone "Contribute" -Exclude DeleteListItems -Connection $connNewSite
}

            # Apply the Home Page Canvas Content
            Write-Verbose -Verbose -Message "Applying EUM-Data-Room-Template home page"
            $canvasContent = Get-PnPFile -Url "$($SiteCollectionRelativeURL)/PnPTemplates/EUM-Data-Room-Template.html" -Connection $connLandingSite -AsString
            
            $eumAdminURLEncoded = $eumAdminURL.Value -replace "https://", "https&#58;//"
            $eumPortalURLEncoded = $eumPortalURL.Value -replace "https://", "https&#58;//"
            $eumAPIApplicationIDURIEncoded = $eumAPIApplicationIDURI.Value -replace "api://", "api&#58;//"
            $eumDataRoomURLEncoded = $siteURL -replace "https://", "https&#58;//"
            $siteTitle = $site.RootWeb.Title
            Write-Output -Verbose -Message "eumAdminUrlEncoded is $($eumAdminURLEncoded)"
            Write-Output -Verbose -Message "siteTitle is $($siteTitle)"
            
            $canvasContent = $canvasContent -replace "~eumAdminURL~", $eumAdminURLEncoded
            $canvasContent = $canvasContent -replace "~eumPortalURL~", $eumPortalURLEncoded
            $canvasContent = $canvasContent -replace "~eumAPIApplicationIDURI~", $eumAPIApplicationIDURIEncoded
            $canvasContent = $canvasContent -replace "~eumDataRoomURL~", $eumDataRoomURLEncoded
            $canvasContent = $canvasContent -replace "~siteTitle~", $siteTitle
            
            $pageNameWithExtension = "Home.aspx"
            $page = Get-PnPListItem -List "SitePages" -Query "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='String'>$($pageNameWithExtension)</Value></Eq></Where></Query></View>" -Connection $connNewSite
            $setPage = Set-PnPListItem -List SitePages -Identity $page.Id -Values @{"CanvasContent1" = $canvasContent } -UpdateType SystemUpdate -Connection $connNewSite            
        }
        elseif (($pnpSiteTemplate -like "*template-partner.xml")) {
            # Connect to the newly created site and get its group ID
            $connNewSite = Helper-Connect-PnPOnline -Url $siteURL
            $site = Get-PnPSite -Connection $connNewSite -Includes GroupId

            # Make sure we have access to the Azure Automation variables
            if (![string]::IsNullOrWhiteSpace($site.GroupId) -and $AzureAutomation) {
                # Get and connect to the EUM Config site
                $eumConfigSite = Get-AutomationVariable -Name 'EUMConfigSiteURL'
                $connEUMConfigSite = Helper-Connect-PnPOnline -Url $eumConfigSite

                #Build the empty JSON group file
                $stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes("{`"Members`": `"`"}"))
                
                # Get the EUM page templates for the partner member page
                $partnerTemplate = Get-PnPListItem -List PublisherPageTemplates -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Group Member Partner Page</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # Get the EUM values needed for webpart
                $eumAdminURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>AdminURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumPortalURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>PortalURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumAPIApplicationIDURI = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>EUM-API-Application-ID-URI</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # EUM page templates found
                if ($partnerTemplate -ne $null -and $partnerTemplate.Id -ne $null) {
                    $folderName = $groupName.Replace(" ", "-").ToLower()
                    # Convert the Azure AD Group into an EUM group
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false; MemberPageTemplate = $partnerTemplate.Id; MemberPage = "Auto-generated URL"; MemberPageURL = "/members/$($folderName)" }
                }
                # No EUM page templates
                else {
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false }
                }
            }   
        }
         elseif (($pnpSiteTemplate -like "*template-client.xml")) {
            # Connect to the newly created site and get its group ID
            $connNewSite = Helper-Connect-PnPOnline -Url $siteURL
            $site = Get-PnPSite -Connection $connNewSite -Includes GroupId

            # Make sure we have access to the Azure Automation variables
            if (![string]::IsNullOrWhiteSpace($site.GroupId) -and $AzureAutomation) {
                # Get and connect to the EUM Config site
                $eumConfigSite = Get-AutomationVariable -Name 'EUMConfigSiteURL'
                $connEUMConfigSite = Helper-Connect-PnPOnline -Url $eumConfigSite

                #Build the empty JSON group file
                $stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes("{`"Members`": `"`"}"))
                
                # Get the EUM page templates for the partner member page
                $clientTemplate = Get-PnPListItem -List PublisherPageTemplates -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Group Member Client Page</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # Get the EUM values needed for webpart
                $eumAdminURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>AdminURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumPortalURL = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>PortalURL</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite
                $eumAPIApplicationIDURI = Get-PnPListItem -List "Suite Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>EUM-API-Application-ID-URI</Value></Eq></Where></Query></View>" -Connection $connEUMConfigSite

                # EUM page templates found
                if ($clientTemplate -ne $null -and $clientTemplate.Id -ne $null) {
                    $folderName = $groupName.Replace(" ", "-").ToLower()
                    # Convert the Azure AD Group into an EUM group
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false; MemberPageTemplate = $clientTemplate.Id; MemberPage = "Auto-generated URL"; MemberPageURL = "/members/$($folderName)" }
                }
                # No EUM page templates
                else {
                    Add-PnPFile -FileName "$($site.GroupId).json" -Folder "UserManagerGroups" -Stream $stream -Connection $connEUMConfigSite -Values @{Title = "$($site.GroupId)"; EUMDisplayName = $groupName; Approvals = $false; SharePointGroup = "$($groupName) Visitors"; SharePointSiteCollectionURL = $siteURL; MemberPageEnabled = $true; GroupWelcomeEmailEnabled = $false }
                }
            }   
        }
        else {
            Write-Verbose "No customizations to apply"
        }
    }

    return $True
}