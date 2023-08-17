using Azure.Core;
using Azure.Identity;
using Microsoft.Graph.Beta;
using Microsoft.Graph.Beta.Models;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // App-ony auth token credential
    private static ClientSecretCredential? _clientSecretCredential;
    // Client configured with app-only authentication
    private static GraphServiceClient? _appClient;

    public static void InitializeGraphForAppOnlyAuth(Settings settings)
    {
        _settings = settings;

        // Ensure settings isn't null
        _ = settings ??
            throw new System.NullReferenceException("Settings cannot be null");

        _settings = settings;

        if (_clientSecretCredential == null)
        {
            _clientSecretCredential = new ClientSecretCredential(
                _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
        }

        if (_appClient == null)
        {
            _appClient = new GraphServiceClient(_clientSecretCredential,
                // Use the default scope, which will request the scopes
                // configured on the app registration
                new[] {"https://graph.microsoft.com/.default"});
        }
    }

    public static async Task<string> GetAppOnlyTokenAsync()
    {
        // Ensure credential isn't null
        _ = _clientSecretCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        // Request token with given scopes
        var context = new TokenRequestContext(new[] {"https://graph.microsoft.com/.default"});
        var response = await _clientSecretCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<SiteCollectionResponse?> GetAllSites()
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Sites.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "webUrl", "id" };
            // Get at most 25 results
            config.QueryParameters.Top = 25;
            // Sort by display name
            config.QueryParameters.Orderby = new[] { "webUrl" };
        });
    }

    public static async Task<BaseSitePageCollectionResponse?> GetAllSitePages(Site site)
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        try 
        {

            var baseSitePages = await _appClient.Sites[site.Id].Pages.GetAsync((config) =>
            {
                // Only request specific properties
                config.QueryParameters.Select = new[] { "title", "id" };
            });

            return baseSitePages;
            
        } 
        catch 
        {
            Console.WriteLine($"Error getting pages for site {site.WebUrl}");
           return new BaseSitePageCollectionResponse();
        }
        
    }
}