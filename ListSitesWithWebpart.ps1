##############################################
# ListSitesWithWebPart.ps1
# Alex Grover - alexgrover@microsoft.com
#
#
##############################################
# Dependencies
##############################################

## Requires the following modules:
try {
    Import-Module Microsoft.Graph.Sites
    Import-Module Microsoft.Graph.Groups
    Import-Module PnP.PowerShell
}
catch {
    Write-Error "Error importing modules required modules - $($Error[0].Exception.Message))"
    Exit
}


# Graph Permissions
# Sites.Read.All
# GroupMember.Read.All
# User.Read.All

# SPO Permissions
# Sites.Read.All

##############################################
# Variables
##############################################

# Auth
$clientId = "38acafba-2eb6-4510-848e-070b493ea4dc"
$tenantId = "groverale.onmicrosoft.com"
$thumbprint = "72A385EF67B35E1DFBACA89180B7B3C8F97453D7"

# Title            WebPartId
# -----            ---------
# YouTube          544dd15b-cf3c-441b-96da-004d5a8cea1d
# Twitter          f6fdf4f8-4a24-437b-a127-32e66a5dd9b4
# Stream (Classic) 275c0095-a77e-4f6d-a2a0-6a7626911518    

# WebPartTypeIds to check
$webPartTypeIds = @(
    "544dd15b-cf3c-441b-96da-004d5a8cea1d",
    "f6fdf4f8-4a24-437b-a127-32e66a5dd9b4",
    "275c0095-a77e-4f6d-a2a0-6a7626911518"
)

# Process all sites or only sites in the input file
$allSites = $false

# List of Sites to check (ignore if $allSites = $true)
$inputSitesCSV = "./SiteCollectionsList.txt"

# Log file location (timestamped with script start time)
$timeStamp = Get-Date -Format "yyyyMMddHHmmss"
$logFileLocation = "Output\WebPartLog-$timeStamp.csv"

# Verbose Logging (Inlcudes all sites and pages, not just sites with webparts)
$verbose = $false

##############################################
# Functions
##############################################

function ConnectToMSGraph 
{  
    try{
        
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint

        select-MgProfile -Name "beta"
    }
    catch{
        Write-Host "Error connecting to MS Graph - $($Error[0].Exception.Message)" -ForegroundColor Red
        Exit
    }
}

function ConnectToPnP ($siteUrl){
    try{
        Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint
    }
    catch{
        Write-Host "Error connecting to PnP" -ForegroundColor Red
    }
}

function Get-SitePages($site)
{
    try {
        $pages = Get-MgSitePage -SiteId $site.Id -Select "id,title,webUrl,name" -ErrorAction Stop
        return $pages
    }
    catch {

        if($Error[0].Exception.Message.Contains("Item not found"))
        {
            Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl "" -pageTitle "" -webPartId "" -type "Warning" -message "Site has no pages library - likely classic/legacy site"
            Write-Host " Error getting pages for $($site.WebUrl), - Site has no pages library - likely classic/legacy site" -ForegroundColor Yellow
            return
        }

        if ($Error[0].Exception.Message.Contains("Access to this site has been blocked"))
        {
            Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl "" -pageTitle "" -webPartId "" -type "Warning" -message "Site has been locked, unable to get site data"
            Write-Host " Site has been locked - site will likely be disposed of in 3 months" -ForegroundColor Yellow
            return
        }

        Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl "" -pageTitle "" -webPartId "" -type "Error" -message "Error getting pages - $($Error[0].Exception.Message)"
        Write-Host " Error getting pages for $($site.WebUrl), - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}
  
function CheckPageContainsWebPartPnP($site, $page, $ownerEmails)
{
    Write-Host "  Checking page (PnP) ($($page.WebUrl)) for webparts" -ForegroundColor White

    try {
        $webParts = Get-PnPPageComponent -Page $page.Name | Select-Object Title, WebPartId, InstanceId -ErrorAction Stop

        foreach ($webPart in $webParts)
        {
            if ($webPartTypeIds.Contains($webPart.WebPartId))
            {
                # Write to Log File
                Write-Host "   Found Webpart:" $webPart.Title -ForegroundColor Green
                foreach ($ownerEmail in $ownerEmails)
                {
                    Write-LogEntry -siteUrl $site.WebUrl -ownerEmail $ownerEmail -pageUrl $page.WebUrl -pageTitle $page.Title -webPartId $webPart.InstanceId -webPartType $webPart.WebPartId -webPartTitle $webPart.Title -type "Success" -message ""
                }
            }
            else 
            {
                if ($verbose)
                {
                    Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl $page.WebUrl -pageTitle $page.Title -type "Info" -message "No target webparts found on page"
                }
            }
        }
    }
    catch {
        Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl $page.WebUrl -pageTitle $pageTitle -webPartId "" -type "Error" -message "Error getting webparts (PnP) - $($Error[0].Exception.Message)"
        Write-Host " Error getting webparts for page on site $($site.WebUrl) - $($Error[0].Exception.Message)" -ForegroundColor Red

        Write-Host "-SiteId $($site.Id) -SitePageId $($page.Id)"
    }
    
}

function Get-SitePageWebparts($site, $page, $ownerEmails)
{
    try {
        Write-Host " Getting webparts for page, $($page.Title) on site $($site.WebUrl)"
        
        $page = Get-MgSitePage -SiteId $site.Id -SitePageId $page.Id -ExpandProperty "webparts" -ErrorAction Stop
        #$webparts = Get-MgBetaSitePageWebPart -SiteId $site.Id -SitePageId $page.Id -ErrorAction Stop
      
        return $page
    }
    catch {
        
        ## all errors are now reprocessed via PnP
        ##if ($Error[0].Exception.Message.Contains("One of the provided arguments is not acceptable"))
        
        # Connect to PnP
        ConnectToPnP -siteUrl $site.WebUrl
        
        # Use PnP to get webparts
        CheckPageContainsWebPartPnP -site $site -page $page -ownerEmails $ownerEmails
        return "pnp"
        

        # Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl $page.WebUrl -pageTitle $pageTitle -webPartId "" -type "Error" -message "Error getting webparts - $($Error[0].Exception.Message)"
        # Write-Host " Error getting webparts for page on site $($site.WebUrl) - $($Error[0].Exception.Message)" -ForegroundColor Red

        # Write-Host "-SiteId $($site.Id) -SitePageId $($page.Id)"
        
    }
}

function Does-PageContainIdentifiedWebparts($siteWebUrl, $page, $outputObjs, $ownerEmails)
{
    Write-Host "  Checking page ($($page.WebUrl)) for webparts" -ForegroundColor White

    if ($page.WebParts.Count -eq 0)
    {
        Write-Host "   No webparts found on page $($page.WebUrl)" -ForegroundColor Yellow
        if ($verbose)
        {
            Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl $page.WebUrl -pageTitle $page.Title -webPartId "" -type "Info" -message "No webparts found on page"
        }
        return
    }

    foreach ($webPart in $page.WebParts)
    {
        if ($webPartTypeIds.Contains($webPart.AdditionalProperties.webPartType))
        {
            # Write to Log File
            Write-Host "   Found Webpart:" $webPart.AdditionalProperties.webPartType -ForegroundColor Green

            foreach ($email in $ownerEmails)
            {
                Write-LogEntry -siteUrl $siteWebUrl -pageUrl $page.WebUrl -pageTitle $page.Title -webPartId $webPart.Id -webPartType $webPart.AdditionalProperties.webPartType -webPartTitle $webPart.AdditionalProperties.data.title -type "Success" -message "" -ownerEmail $email
            }
        }
        else 
        {
            if ($verbose)
            {
                Write-LogEntry -siteUrl $siteWebUrl -ownerEmail "" -pageUrl $page.WebUrl -pageTitle $page.Title -type "Info" -message "No target webparts found on page"
            }
        }
    }
}

function ReadSitesFromTxtFile($siteListCSVFile) {
    $siteList = Get-Content $siteListCSVFile
    return $siteList
}

function Get-Sites
{
    try {
        
        if (!$allSites) {
            $siteList = ReadSitesFromTxtFile($inputSitesCSV)
            $sites = Get-MgSite -Property "siteCollection,webUrl,id" -All | Where-Object { !($_.WebUrl.Contains("my.sharepoint.com"))} | where { $siteList -contains $_.WebUrl } -ErrorAction Stop
            return $sites 
        }

        # Get all sites, filter out OneDrive sites
        $sites = Get-MgSite -Property "siteCollection,webUrl,id" -All | Where-Object { !($_.WebUrl.Contains("my.sharepoint.com"))} -ErrorAction Stop
        return $sites #| where {$_.WebUrl.Contains("/home")}
    }
    catch {
        Write-Host " Error getting sites" -ForegroundColor Red
    }   
}

function Get-AllSubsites ($site, $subsites)
{
    Write-Host " Getting Subsites for: $($site.webUrl)..."

    # Add the site to the subsites array
    $subsites.Add($site) | Out-Null

    try {
        # Get the site's children
        $children = Get-MgSubSite -Property "siteCollection,webUrl,id" -SiteId $site.Id -All -ErrorAction Stop

        # Recursively get all subsites and their descendants
        foreach ($child in $children) {

            # Recursively get the subsite's descendants
            Get-AllSubsites -site $child -subsites $subsites
        }
    }
    catch {
        if ($Error[0].Exception.Message.Contains("Access to this site has been blocked"))
        {
            # Swallow the error - will be caught in next function
            return
        }
    } 
}

function Write-LogEntry($siteUrl, $pageUrl, $pageTitle, $webPartId, $webPartType, $type, $message, $webPartTitle, $ownerEmail)
{
    $logLine = New-Object -TypeName PSObject -Property @{
        Type = $type
        LogTime = Get-Date
        WebParttype = $webPartType
        SiteUrl = $siteUrl
        PageUrl = $pageUrl
        WebPartId = $webPartId
        WebPartTitle = $webPartTitle
        Notes = $message
        OwnerEmail = $ownerEmail
    }

    $logLine | Export-Csv -Path $logFileLocation -NoTypeInformation -Append
}

function ProcessSite($site, $ownerEmails)
{
    ## Get all pages
    $pages = Get-SitePages -site $site

    if ($pages.Count -eq 0)
    {
        Write-Host " No pages found on site $($site.WebUrl)" -ForegroundColor Yellow
        if ($verbose)
        {
            Write-LogEntry -siteUrl $site.WebUrl -ownerEmail "" -pageUrl "" -pageTitle "" -webPartId "" -type "Info" -message "No pages found on site"
        }
        return
    }

    ## Loop through pages
    foreach ($page in $pages)
    {
        ## Get all webparts
        $page = Get-SitePageWebparts -site $site -page $page -ownerEmails $ownerEmails
        
        ## Return if page has been processed using PnP
        if ($page -eq "pnp")
        {
            continue
        }
        ## Check if page contains webpart
        Does-PageContainIdentifiedWebparts -siteWebUrl $site.WebUrl -page $page -outputObjs $outputObjs -ownerEmails $ownerEmails
    }
}

function GetSiteOwner($site)
{
    $owners = @()

    Write-Host "Checking site ($($site.WebUrl)) for owners" -ForegroundColor White

    try {
        $defaultDrive = Get-MgSiteDefaultDrive -SiteId $site.Id -Property Owner

        if (![string]::IsNullOrEmpty($defaultDrive.Owner.User.AdditionalProperties.email))
        {
            Write-Host "USER OWNER FOUND" -ForegroundColor Green
            $admin = $defaultDrive.Owner.User.AdditionalProperties.email
            $owners += $admin

            Write-Host "Found $($admin) as primary owner" -ForegroundColor Green

            # drop in pnp to check if there are mutiple owners
            ConnectToPnP -siteUrl $site.WebUrl
            $siteOwners = Get-PnPGroup -AssociatedOwnerGroup | Get-PnPGroupMember

            foreach ($siteOwner in $siteOwners)
            {
                # We may have Groups, Sec Groups or users

                # User - just add the email
                if ($siteOwner.LoginName.Contains("|membership|"))
                {
                    Write-Host "  Found: $($siteOwner.Email) as an owner" -ForegroundColor Green
                    $owners += $siteOwner.Email
                    continue
                }

                # Group - get the group *Owners*
                # Group membership can only be users, not groups so no need to recurse
                if ($siteOwner.LoginName.Contains("|federateddirectoryclaimprovider|"))
                {
                    Write-Host "  Found: $($siteOwner.Email) (Group) as an owner" -ForegroundColor Green
                    $groupId = $siteOwner.LoginName.Split("|")[2]

                    Write-Host "  Getting Group Members" -ForegroundColor White
                    ## Seems to only return ids, so we need to get the users
                    ## Instances where we the group id isn't properly formed
                    $members = Get-MgGroupOwner -GroupId $groupId.Substring(0, [Math]::Min($groupId.Length, 36)) -Property "userPrincipalName" 

                    foreach ($member in $members)
                    {
                        Write-Host "   Found: $($member.AdditionalProperties.userPrincipalName) as an infered owner" -ForegroundColor Green
                        $owners += $member.AdditionalProperties.userPrincipalName
                    }
                    continue
                }

                # Sec Group - get the group *Members* as it's members who are owners of the site
                # Sec Group membership can be all sorts so we need to recurse
                if ($siteOwner.LoginName.Contains("|tenant|"))
                {
                    Write-Host "  Found: $($siteOwner.LoginName) (SecGroup) as an owner" -ForegroundColor Green
                    $groupId = $siteOwner.LoginName.Split("|")[2]

                    $members = Get-MgGroupMember -GroupId $groupId -Property "userPrincipalName,id" 

                    foreach ($member in $members)
                    {
                        ## TODO - Check we have a user and not a group (if group we need to get memebers)
                        if ($null -ne $member.AdditionalProperties.userPrincipalName)
                        {
                            Write-Host "   Found: $($member.AdditionalProperties.userPrincipalName) as an infered owner" -ForegroundColor Green
                            ## We have a user
                            $owners += $member.AdditionalProperties.userPrincipalName
                            continue
                        }
                        else {
                            Write-Host "    Found: $($member.Id) (group) as an infered owner" -ForegroundColor Green
                            Write-Host "     Getting Group Members" -ForegroundColor White
                            ## We have a group and must get the members of that group
                            $users = @()
                            $owners += GetSecGroupMembers -groupId $member.Id -users $users
                        }
                    }
                    continue
                }
            }
        }
        else
        {
            Write-Host "GROUP OWNER FOUND" -ForegroundColor Green

            $group = $defaultDrive.Owner.AdditionalProperties.group
            Write-Host "Found $($group.email) (Group) as primary owner" -ForegroundColor Green
            ## Get the group Owners
            $groupOwners = Get-MgGroupOwner -GroupId $group.id -Property "userPrincipalName"
            Write-Host " Getting Group Owners" -ForegroundColor Green

            foreach ($groupOwner in $groupOwners)
            {
                Write-Host "  Found: $($groupOwner.AdditionalProperties.userPrincipalName) as an owner" -ForegroundColor Green
                $owners += $groupOwner.AdditionalProperties.userPrincipalName
            }

        }
        # Return unique list of owners (emails)
        return $owners | Get-Unique
    }
    catch {
        Write-Host " Error getting site owner for $($site.WebUrl) - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

## Recursive function to get members (users) of a security group
function GetSecGroupMembers($groupId, $users)
{
    $members = Get-MgGroupMember -GroupId $groupId -Property "userPrincipalName,id"

    foreach ($member in $members)
    {
        if ($null -ne $member.AdditionalProperties.userPrincipalName)
        {
            Write-Host "     Found Member - $($member.AdditionalProperties.userPrincipalName) " -ForegroundColor Green
            ## We have a user
            $users += $member.AdditionalProperties.userPrincipalName
        }
        else {
            Write-Host "     Found Group - $($member.Id) " -ForegroundColor Green
            ## We have a group and must get the members (that are users) of that group
            $users += GetSecGroupMembers -groupId $member.Id -users $users
        }
    }
    return $users
}

##############################################
# Main
##############################################

## Connect to Mdestinationraph
ConnectToMSGraph

## Get all sites
$sites = Get-Sites 

## hold the log entries
$outputObjs = @()

## Clear the CSV
$outputObjs | Export-Csv -Path $logFileLocation -NoTypeInformation -Force

## Loop through sites
foreach ($site in $sites)
{

    # Get Owner Details (SiteCollection)
    $ownerEmails = GetSiteOwner -site $site

    # Get all subsites and their descendants
    # Root site is added to the array of sites
    $subsites = New-Object System.Collections.ArrayList
    Get-AllSubsites -site $site -subsites $subsites

    foreach($subsite in $subsites)
    {
        ## Process Sites
        ProcessSite -site $subsite -ownerEmails $ownerEmails
    }
}

