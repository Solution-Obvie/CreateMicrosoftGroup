<#
.\deploy.ps1 -TenantName "o365testtenantobvie" -TenantID "7f894c83-1d00-4ead-8d45-010c83badee5" -RequestsSiteName "Request a group site" -RequestsSiteDesc "Used to store Teams Requests" -ManagedPath "sites" -AppName "Requestagroupsite" 
#>



Param(
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true)]
    [String]
    $TenantName,

    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true)]
    [String]
    $TenantId,

    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true)]
    [String]
    $RequestsSiteName,

    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true)]
    [String]
    $RequestsSiteDesc,

    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true)]
    [String]
    $ManagedPath,

    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true)]
    [String]
    $AppName

)

Add-Type -AssemblyName System.Web

if (Get-InstalledModule  -Name "microsoft.online.sharepoint.powershell") {
Write-Host "SharePoint Online Powershell installed"
}
else{
Install-Module microsoft.online.sharepoint.powershell -Scope CurrentUser
}

if (Get-InstalledModule  -Name "SharePointPnPPowerShellOnline") {
Write-Host "SharePoint PnP Powershell Online installed"
}
else{
Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser -MinimumVersion 3.19.2003.0 -Force
}

if (Get-InstalledModule  -Name "ImportExcel") {
Write-Host "Import Excel installed"
}
else{
Install-Module ImportExcel -Scope CurrentUser
}

if (Get-InstalledModule  -Name "Az") {
Write-Host "Az installed"
}
else{
Install-Module Az -AllowClobber -Scope CurrentUser
}

if (Get-InstalledModule  -Name "AzureAD") {
Write-Host "Azure AD installed"
}
else{
Install-Module AzureAD -Scope CurrentUser
}

if (Get-InstalledModule  -Name "WriteAscii") {
    Write-host "Write Ascii installed"
}
else{
    Install-Module WriteAscii -Scope CurrentUser
}



# Variables
$packageRootPath = "..\"
$templatePath = "Templates\teamsautomate-sitetemplate.xml"
$settingsPath = "Scripts\Settings\SharePoint List items.xlsx"

#  Worksheet
$siteRequestSettingsWorksheetName = "Request Settings"
$teamsTemplatesWorksheetName = "Teams Templates"


#  lists
$requestsListName = "Group Requests"
$requestSettingsListName = "Group Request Settings"
$teamsTemplatesListName = "Teams Templates"

#  Field names
$TitleFieldName = "Title"
$TeamNameFieldName = "Group Name"

$tenantUrl = "https://$tenantName.sharepoint.com"
$tenantAdminUrl = "https://$tenantName-admin.sharepoint.com"

# Remove any spaces in the site name to create the alias
$requestsSiteAlias = $RequestsSiteName -replace (' ', '')
$requestsSiteUrl = "https://$tenantName.sharepoint.com/$ManagedPath/$requestsSiteAlias"

# Global variables
$global:context = $null
$global:requestsListId = $null
$global:teamsTemplatesListId = $null
$global:appId = $null
$global:appSecret = $null
$global:appServicePrincipalId = $null
$global:siteClassifications = $null
$global:location = $null


# Create site and apply provisioning template
function CreateRequestsSharePointSite {
     try {

        Write-Host "### TEAMS REQUESTS SITE CREATION ###`nCreating Teams Requests SharePoint site..." -ForegroundColor Yellow

    $site = Get-PnPTenantSite -Url $requestsSiteUrl -ErrorAction SilentlyContinue

    if (!$site) {

        # Site will be created with current user connected to PnP as the owner/primary admin
        New-PnPSite -Type TeamSite -Title $RequestsSiteName -Alias $requestsSiteAlias -Description $RequestsSiteDesc

        Write-Host "Site created`n**TEAMS REQUESTS SITE CREATION COMPLETE**" -ForegroundColor Green
    }         

    else {
        Write-Host "Site already exists! Do you wish to overwrite?" -ForegroundColor Red
        $overwrite = Read-Host " ( y (overwrite) / n (exit) )"
        if ($overwrite -ne "y") {
            break
        }
            
    }
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Host "Error occured while creating of the SharePoint site: $errorMessage" -ForegroundColor Red
}
}

# Configure the new site
function ConfigureSharePointSite {

    try {

        Write-Host "### REQUESTS SPO SITE CONFIGURATION ###`nConfiguring SharePoint site..." -ForegroundColor Yellow

        Write-Host "Applying provisioning template..." -ForegroundColor Yellow

        Apply-PnPProvisioningTemplate -Path (Join-Path $packageRootPath $templatePath) -ClearNavigation

        Write-Host "Applied template" -ForegroundColor Green

        $context = Get-PnPContext
        # Ensure Site Assets
        $web = $context.Web
        $context.Load($web)
        $context.Load($web.Lists)
        $context.ExecuteQuery()

        # Rename Title field
        $siteRequestsList = Get-PnPList $requestsListName
        $global:requestsListId = $siteRequestsList.Id
        $fields = $siteRequestsList.Fields
        $context.Load($fields)
        $context.ExecuteQuery()

        $titleField = $fields | Where-Object { $_.InternalName -eq $TitleFieldName }
        $titleField.Title = $TeamNameFieldName
        $titleField.UpdateAndPushChanges($true)
        $context.ExecuteQuery()

        # Adding settings in Site request Settings list
        $siteRequestsSettingsList = Get-PnPList $requestSettingsListName
        $context.Load($siteRequestsSettingsList)
        $context.ExecuteQuery()

        $siteRequestSettings = Import-Excel "$packageRootPath$settingsPath" -WorksheetName $siteRequestSettingsWorksheetName
        foreach ($setting in $siteRequestSettings) {
            if ($setting.Title -eq "TenantURL") {
                $setting.Value = $tenantUrl
            }
            if ($setting.Title -eq "SPOManagedPath") {
                $setting.Value = $ManagedPath
            }
             if ($setting.Title -eq "AppID") {
                $setting.Value = $global:appId
            }
              if ($setting.Title -eq "AppSecret") {
                $setting.Value = $global:appSecret
            }
                if ($setting.Title -eq "SharePointURL") {
                $setting.Value = $requestsSiteUrl
            }
              if ($setting.Title -eq "TenantID") {
                $setting.Value = $TenantId
            }
            $listItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
            $newItem = $siteRequestsSettingsList.AddItem($listItemCreationInformation)
            $newitem["Title"] = $setting.Title
            $newitem["Description"] = $setting.Description
            # Hide site classifications option in Power App if no site classifications were found in the tenant
            $newitem["Value"] = $setting.Value   
            $newitem.Update()
            $context.ExecuteQuery()
        }

        Write-Host "Added settings to Site Requests Settings list" -ForegroundColor Green

        # Adding templates to Teams Templates list
        $teamsTemplatesList = Get-PnPList $teamsTemplatesListName
        $context.Load($teamsTemplatesList)
        $context.ExecuteQuery()
        $global:teamsTemplatesListId = $teamsTemplatesList.Id

        $teamsTemplates = Import-Excel "$packageRootPath$settingsPath" -WorksheetName $teamsTemplatesWorksheetName
        foreach ($template in $teamsTemplates) {
            If (!$isEdu -and ($template.BaseTemplateId -eq "educationStaff" -or $template.BaseTemplateId -eq "educationProfessionalLearningCommunity")) {
                # Tenant is not an EDU tenant  - do nothing
            }
            else {
                $listItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                $newItem = $teamsTemplatesList.AddItem($listItemCreationInformation)
                $newItem["Title"] = $template.Title
                $newItem["BaseTemplateType"] = $template.BaseTemplateType
                $newItem["BaseTemplateId"] = $template.BaseTemplateId
                $newItem["Description"] = $template.Description
                $newItem["FirstPartyTemplate"] = $template.FirstPartyTemplate
                $newitem.Update()
                $context.ExecuteQuery()
            }
        }
        Write-Host "Added templates to Teams Templates list" -ForegroundColor Green

        Write-Host "Finished configuring site" -ForegroundColor Green

    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error occured while configuring the SharePoint site: $errorMessage" -ForegroundColor Red
    }
}

function GetAzureADApp {
    param ($appName)

    $app = az ad app list --filter "displayName eq '$appName'" | ConvertFrom-Json

    return $app

}

function CreateAzureADApp {
  
    try {
        Write-Host "### AZURE AD APP CREATION ###" -ForegroundColor Yellow

        # Check if the app already exists - script has been previously executed
        $app = GetAzureADApp $appName

        if (-not ([string]::IsNullOrEmpty($app))) {

           
            # Update azure ad app registration using CLI
            Write-Host "Azure AD App '$appName' already exists - updating existing app..." -ForegroundColor Yellow

            az ad app update --id $app.appId --required-resource-accesses './manifest.json' --password $global:appSecret 

            Write-Host "Waiting for app to finish updating..."

            Start-Sleep -s 60

            Write-Host "Updated Azure AD App" -ForegroundColor Green

        } 
        else {
            # Create the app
            Write-Host "Creating Azure AD App - '$appName'..." -ForegroundColor Yellow

            # Create azure ad app registration using CLI
            az ad app create --display-name $appName --required-resource-accesses './manifest.json' --password $global:appSecret --end-date '2299-12-31T11:59:59+00:00' 

            Write-Host "Waiting for app to finish creating..."

            Start-Sleep -s 60
            
            Write-Host "Created Azure AD App" -ForegroundColor Green

        }

        $app = GetAzureADApp $appName
        $global:appId = $app.appId

        Write-Host "Granting admin consent for Microsoft Graph..." -ForegroundColor Yellow

        # Grant admin consent for app registration required permissions using CLI
        az ad app permission admin-consent --id $global:appId
        
        Write-Host "Waiting for admin consent to finish..."

        Start-Sleep -s 60
        
        Write-Host "Granted admin consent" -ForegroundColor Green

        # Get service principal id for the app we created
        $global:appServicePrincipalId = Get-AzADServicePrincipal -DisplayName $appName | Select-Object -ExpandProperty Id

        Write-Host "### AZURE AD APP CREATION FINISHED ###" -ForegroundColor Green
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error occured while creating an Azure AD App: $errorMessage" -ForegroundColor Red
    }
}


Write-Host "### DEPLOYMENT SCRIPT STARTED ###" -ForegroundColor Magenta

$guid = New-Guid

$global:appSecret = ([System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($guid))))
$global:encodedAppSecret = [System.Web.HttpUtility]::UrlEncode($global:appSecret) 

# Initialise connections - Azure Az/CLI
Write-Host "Launching Azure sign-in..." -ForegroundColor Yellow
#$azConnect = Connect-AzAccount -Subscription $SubscriptionId -Tenant $TenantId
Write-Host "Launching Azure AD sign-in..." -ForegroundColor Yellow
Connect-AzureAD
Write-Host "Launching Azure CLI sign-in..." -ForegroundColor Yellow
$cliLogin = az login --allow-no-subscriptions
Write-Host "Connected to Azure" -ForegroundColor Green
# Connect to PnP
Write-Host "Launching PnP sign-in..." -ForegroundColor Yellow
$pnpConnect = Connect-PnPOnline -Url $tenantAdminUrl -Credentials (Get-Credential)
#$pnpConnect = Connect-PnPOnline -Url $tenantAdminUrl -UseWebLogin
Write-Host "Connected to SPO" -ForegroundColor Green
CreateAzureADApp
CreateRequestsSharePointSite
# Connect to the new site
$pnpConnect = Connect-PnPOnline $requestsSiteUrl -UseWebLogin
ConfigureSharePointSite
Write-Host "DEPLOYMENT COMPLETED SUCCESSFULLY" -ForegroundColor Green