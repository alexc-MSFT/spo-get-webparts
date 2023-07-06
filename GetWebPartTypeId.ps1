# Connect to SharePoint Online
Connect-PnPOnline -Url "https://groverale.sharepoint.com/sites/home" -Interactive

# Specify the page Name and the Web Part title
# https://groverale.sharepoint.com/sites/home/SitePages/FAQs.aspx
$pageName = "FAQs"
$webPartTitle = "Stream"

# Get the page and the Web Part
$page = Get-PnPClientSidePage -Identity $pageName
$webPart = Get-PnPPageComponent -Page $page | where { $_.Title -eq $webPartTitle }

# Get the Web Part Type ID
$webPartTypeId = $webPart.WebPartId
Write-Host "Web Part Type ID for $webPartTitle : $webPartTypeId"
