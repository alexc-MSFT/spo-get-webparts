# ListSitesWithWebPart.ps1

This PowerShell script is designed to list SharePoint sites and pages that contain specific web parts. It connects to the Microsoft Graph API and SharePoint Online using the PnP PowerShell module to retrieve site and page information. Site Owner details are also returned from the script.

## Dependencies
The script requires the following PowerShell modules:
- Microsoft.Graph.Sites
- Microsoft.Graph.Groups
- PnP.PowerShell

Please make sure these modules are installed before running the script. If any of the modules are missing, the script will display an error and exit.

## Graph Permissions
The script requires the following Microsoft Graph permissions:
- Sites.Read.All
- GroupMember.Read.All
- User.Read.All

## SPO Permissions
The script requires the following SharePoint Online permission:
- Sites.Read.All

## App Registration

The script requires an app registration in Azure AD to authenticate with Microsoft Graph and SharePoint Online. The app registration allows the script to access the required resources and perform the necessary operations. To set up the app registration, follow these steps:

1. Go to the Azure portal (portal.azure.com) and navigate to the Azure Active Directory section.
2. Select "App registrations" and click on "New registration" to create a new app registration.
3. Provide a name for the app registration and choose the appropriate supported account types (e.g., single tenant, multi-tenant).
4. Once the app registration is created, note down the "Client ID" (also known as the Application ID) as it will be used as the `$clientId` in the script.

## Creating a Self-Signed Certificate

The script requires a self-signed certificate that is used for authentication. Follow these steps to create a self-signed certificate using the script:

1. Run the following code to create a self-signed certificate (Note. This script needs to be run from the device that will run the script):

   ```powershell
   $certname = "StreamWebPartCert"
   $cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
   Export-Certificate -Cert $cert -FilePath "$certname.cer"  ## Specify your preferred location
    ```

After running the code, the certificate will be saved with the .cer extension in the specified location.

The .cer file will need to be uploaded to the App Registration.

## Variables
The script includes several variables that can be customized according to your needs:

- `$clientId`: The client ID for authenticating with Microsoft Graph.
- `$tenantId`: The ID of your Azure AD tenant.
- `$thumbprint`: The thumbprint of the certificate used for authentication.
- `$webPartTypeIds`: An array of web part type IDs to check for on the pages.
- `$allSites`: Set to `$true` if you want to process all sites or `$false` to process only sites listed in the input file.
- `$inputSitesCSV`: The path to the input file that contains a list of sites to process. This parameter is ignored if `$allSites` is set to `$true`.
- `$logFileLocation`: The location and name of the log file. The log file will be timestamped with the script start time.
- `$verbose`: Set to `$true` to enable verbose logging, which includes all sites and pages, not just sites with web parts.

You can modify these variables to fit your environment and requirements.

### $webPartTypeIds

This is the array of webpart Ids that you are looking for.

You can use the following script to output the webpart id for a known webpart

```powershell
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
```

## Output

The script has two output modes, verbose and normal. Verbose will log a line for every page / webpart checked (`Info`). Normal will only log `Errors`, `Warnings` and `Success` it does not log `Info`. A line for every owner of a site is added for every webpart found. This has been done to make the sending of notifications simpler. There will be a line item for every webpart found. i.e. if a page has two webparts there will be a line for each site owner twice for that page. If a site has 3 pages with the webpart there will be 3 lines for each site owner. A sample csv can be seen below

|SiteUrl                                                 |Type   |OwnerEmail                                             |LogTime            |WebPartTitle|Notes|WebPartId                           |PageUrl                         |WebParttype                         |
|--------------------------------------------------------|-------|-------------------------------------------------------|-------------------|------------|-----|------------------------------------|--------------------------------|------------------------------------|
|https://groverale.sharepoint.com/sites/home             |Success|alex@groverale.onmicrosoft.com                         |07/07/2023 09:53:39|YouTube     |     |1f3ddebe-b649-4398-9f09-be8a4c1102d4|SitePages/FAQs.aspx             |544dd15b-cf3c-441b-96da-004d5a8cea1d|
|https://groverale.sharepoint.com/sites/home             |Success|serviceaccount@groverale.onmicrosoft.com               |07/07/2023 09:53:40|YouTube     |     |1f3ddebe-b649-4398-9f09-be8a4c1102d4|SitePages/FAQs.aspx             |544dd15b-cf3c-441b-96da-004d5a8cea1d|
|https://groverale.sharepoint.com/sites/home             |Success|alex@groverale.onmicrosoft.com                         |07/07/2023 09:53:40|Twitter     |     |490ff3fa-d3d8-4c7a-9c09-97b20233ee6e|SitePages/FAQs.aspx             |f6fdf4f8-4a24-437b-a127-32e66a5dd9b4|
|https://groverale.sharepoint.com/sites/home             |Success|serviceaccount@groverale.onmicrosoft.com               |07/07/2023 09:53:40|Twitter     |     |490ff3fa-d3d8-4c7a-9c09-97b20233ee6e|SitePages/FAQs.aspx             |f6fdf4f8-4a24-437b-a127-32e66a5dd9b4|
|https://groverale.sharepoint.com/sites/home             |Success|alex@groverale.onmicrosoft.com                         |07/07/2023 09:53:40|Stream      |     |e943ba8b-cf30-4e05-94e7-09772c9e1be2|SitePages/FAQs.aspx             |275c0095-a77e-4f6d-a2a0-6a7626911518|
|https://groverale.sharepoint.com/sites/home             |Success|serviceaccount@groverale.onmicrosoft.com               |07/07/2023 09:53:40|Stream      |     |e943ba8b-cf30-4e05-94e7-09772c9e1be2|SitePages/FAQs.aspx             |275c0095-a77e-4f6d-a2a0-6a7626911518|
|https://groverale.sharepoint.com/sites/home             |Success|alex@groverale.onmicrosoft.com                         |07/07/2023 09:53:40|YouTube     |     |d954e9ea-b489-44f4-9f6f-14270918387c|SitePages/Let's-go-collapse.aspx|544dd15b-cf3c-441b-96da-004d5a8cea1d|
|https://groverale.sharepoint.com/sites/home             |Success|serviceaccount@groverale.onmicrosoft.com               |07/07/2023 09:53:40|YouTube     |     |d954e9ea-b489-44f4-9f6f-14270918387c|SitePages/Let's-go-collapse.aspx|544dd15b-cf3c-441b-96da-004d5a8cea1d|
|https://groverale.sharepoint.com/sites/home/none/subtwit|Success|alex@groverale.onmicrosoft.com                         |07/07/2023 09:53:41|Twitter     |     |177f42d5-79ef-476f-a7a8-9ba5d9e02869|SitePages/Home.aspx             |f6fdf4f8-4a24-437b-a127-32e66a5dd9b4|
|https://groverale.sharepoint.com/sites/home/none/subtwit|Success|serviceaccount@groverale.onmicrosoft.com               |07/07/2023 09:53:41|Twitter     |     |177f42d5-79ef-476f-a7a8-9ba5d9e02869|SitePages/Home.aspx             |f6fdf4f8-4a24-437b-a127-32e66a5dd9b4|
|https://groverale.sharepoint.com/sites/home/twitter     |Success|alex@groverale.onmicrosoft.com                         |07/07/2023 09:53:42|Twitter     |     |cd9799c6-8a81-4d8c-b92d-fb08b62cdb5f|SitePages/Home.aspx             |f6fdf4f8-4a24-437b-a127-32e66a5dd9b4|
|https://groverale.sharepoint.com/sites/home/twitter     |Success|serviceaccount@groverale.onmicrosoft.com               |07/07/2023 09:53:42|Twitter     |     |cd9799c6-8a81-4d8c-b92d-fb08b62cdb5f|SitePages/Home.aspx             |f6fdf4f8-4a24-437b-a127-32e66a5dd9b4|
|https://groverale.sharepoint.com/sites/home/youtube     |Success|alex@groverale.onmicrosoft.com                         |07/07/2023 09:53:43|YouTube     |     |345e1bdf-7ef1-4cd7-b14e-c62abedc6ff3|SitePages/Home.aspx             |544dd15b-cf3c-441b-96da-004d5a8cea1d|
|https://groverale.sharepoint.com/sites/home/youtube     |Success|serviceaccount@groverale.onmicrosoft.com               |07/07/2023 09:53:43|YouTube     |     |345e1bdf-7ef1-4cd7-b14e-c62abedc6ff3|SitePages/Home.aspx             |544dd15b-cf3c-441b-96da-004d5a8cea1d|


## Script Deepdive

### Functions
The script includes several functions that handle different tasks:

- `ConnectToMSGraph`: Connects to Microsoft Graph using the specified client ID, tenant ID, and certificate thumbprint. This function also selects the beta profile for accessing the beta endpoint.
- `ConnectToPnP`: Connects to SharePoint Online using the specified site URL, client ID, tenant ID, and certificate thumbprint.
- `Get-SitePages`: Retrieves all pages from a site using the Microsoft Graph API. It returns a collection of page objects.
- `CheckPageContainsWebPartPnP`: Checks if a page contains any of the specified web parts using the PnP PowerShell module.
- `Get-SitePageWebparts`: Retrieves all web parts from a page using the Microsoft Graph API. It returns a page object with the web parts included.
- `Does-PageContainIdentifiedWebparts`: Checks if a page contains any of the specified web parts using the web parts retrieved from the Microsoft Graph API.
- `ReadSitesFromTxtFile`: Reads a list of sites from a text file and returns an array of site URLs.
- `Get-Sites`: Retrieves a collection of SharePoint sites using the Microsoft Graph API. It filters out OneDrive sites and optionally filters based on the input file if `$allSites` is set to `$false`.
- `Get-AllSubsites`: Retrieves all subsites and their descendants for a given site recursively.
- `Write-LogEntry`: Writes a log entry to the specified log file with details such as the site URL, page URL, web part ID, and message.

These functions handle various operations within the script and can be customized or extended as needed.

### Main
The main part of the script performs the following steps:

1. Connects to Microsoft Graph using the `ConnectToMSGraph` function.
2. Retrieves all SharePoint sites using the `Get-Sites` function.
3. Loops through each site and performs the following steps:
   - Retrieves the site owner details using the `GetSiteOwner` function.
   - Retrieves all subsites and their descendants using the `Get-AllSubsites` function.
   - Processes each site and subsite using the `ProcessSite` function.
4. Within the `ProcessSite` function, the script retrieves all pages for the site or subsite using the `Get-SitePages` function.
5. Loops through each page and performs the following steps:
   - Retrieves all web parts for the page using the `Get-SitePageWebparts` function.
   - Checks if the page contains any of the specified web parts using the `Does-PageContainIdentifiedWebparts` function.
   - Writes a log entry for each identified web part using the `Write-LogEntry` function.

The script generates a log file with the specified name and location, which contains the results of the web part search.

Make sure to customize the script variables according to your environment and run the script in a PowerShell environment with the required modules and permissions.
