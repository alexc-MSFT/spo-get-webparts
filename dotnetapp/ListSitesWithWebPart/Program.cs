Console.WriteLine(".NET Graph App-only Tutorial\n");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

int choice = -1;

await ListSitesAsync();

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. List webparts (all sites)");
    Console.WriteLine("3. List webparts (selected sites)");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch(choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // List sites
            await ListSitesAsync();
            break;
        case 3:
            // Run any Graph code
            await MakeGraphCallAsync();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}

void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForAppOnlyAuth(settings);
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        var appOnlyToken = await GraphHelper.GetAppOnlyTokenAsync();
        Console.WriteLine($"App-only token: {appOnlyToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting app-only access token: {ex.Message}");
    }
}

async Task ListSitesAsync()
{
    try
    {
        var sitePage = await GraphHelper.GetAllSites();

        if (sitePage?.Value == null)
        {
            Console.WriteLine("No results returned.");
            return;
        }

        // Output each users's details
        foreach (var site in sitePage.Value)
        {
            Console.WriteLine($"Site: {site.WebUrl ?? "NO NAME"}");
            Console.WriteLine($"  ID: {site.Id}");

            // Get all site pages
            var baseSitePagePages = await GraphHelper.GetAllSitePages(site);
            if (baseSitePagePages?.Value == null)
            {
                Console.WriteLine("No pages for this site.");
                continue;
            }

            foreach (var baseSitePage in baseSitePagePages.Value)
            {
                Console.WriteLine($"  Site page: {baseSitePage.Title ?? "NO NAME"}");
                Console.WriteLine($"    ID: {baseSitePage.Id}");
            }
        }

        // If NextPageRequest is not null, there are more users
        // available on the server
        // Access the next page like:
        // var nextPageRequest = new UsersRequestBuilder(userPage.OdataNextLink, _appClient.RequestAdapter);
        // var nextPage = await nextPageRequest.GetAsync();
        var moreAvailable = !string.IsNullOrEmpty(sitePage.OdataNextLink);

        Console.WriteLine($"\nMore sites available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting sites: {ex.Message}");
    }
}

async Task MakeGraphCallAsync()
{
    // TODO
}