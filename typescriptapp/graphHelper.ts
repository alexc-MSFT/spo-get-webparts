import 'isomorphic-fetch';
import { ClientSecretCredential } from '@azure/identity';
import { Client, GraphRequestOptions, PageCollection, PageIterator, PageIteratorCallback } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from
    '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

import { AppSettings } from './appSettings';
import { Site } from '@microsoft/microsoft-graph-types';

let _settings: AppSettings | undefined = undefined;
let _clientSecretCredential: ClientSecretCredential | undefined = undefined;
let _appClient: Client | undefined = undefined;

export function initializeGraphForAppOnlyAuth(settings: AppSettings) {
    // Ensure settings isn't null
    if (!settings) {
        throw new Error('Settings cannot be undefined');
    }

    _settings = settings;

    if (!_clientSecretCredential) {
        _clientSecretCredential = new ClientSecretCredential(
            _settings.tenantId,
            _settings.clientId,
            _settings.clientSecret
        );
    }

    if (!_appClient) {
        const authProvider = new TokenCredentialAuthenticationProvider(_clientSecretCredential, {
            scopes: ['https://graph.microsoft.com/.default']
        });

        _appClient = Client.initWithMiddleware({
            authProvider: authProvider,
            // Use beta endpoint
            defaultVersion: 'beta'
        });
    }
}

export async function getAppOnlyTokenAsync(): Promise<string> {
    // Ensure credential isn't undefined
    if (!_clientSecretCredential) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    // Request token with given scopes
    const response = await _clientSecretCredential.getToken([
        'https://graph.microsoft.com/.default'
    ]);
    return response.token;
}

export async function getAllSitesAsync(): Promise<Site[]> {
    // Ensure client isn't undefined
    if (!_appClient) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    let sites: Site[] = [];

    const response = await _appClient?.api('/sites')
        .select(['webUrl', 'id'])
        .top(50)
        .orderby('webUrl')
        .get();


        const callback: PageIteratorCallback = (site: Site) => {
            console.log(site.webUrl);
            sites.push(site);
            return true;
          };

          // A set of request options to be applied to
        // all subsequent page requests
        const requestOptions: GraphRequestOptions = {
            // // Re-add the header to subsequent requests
            // headers: {
            //     Prefer: 'outlook.body-content-type="text"',
            // },
        };
    
          // Creating a new page iterator instance with client a graph client
        // instance, page collection response from request and callback
        const pageIterator = new PageIterator(
            _appClient,
            response,
            callback,
            requestOptions
        );
    
        

    await pageIterator.iterate();

    return sites;
}

export async function getSitesAsync(): Promise<PageCollection> {
    // Ensure client isn't undefined
    if (!_appClient) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    return _appClient?.api('/sites')
        .select(['webUrl', 'id'])
        .top(50)
        .orderby('webUrl')
        .get();
}

// A callback function to be called for every item in the collection.
// This call back should return boolean indicating whether not to
// continue the iteration process.

  


export async function getSitePagesAsync(site: Site): Promise<PageCollection> {
    // Ensure client isn't undefined
    if (!_appClient) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    return _appClient?.api(`/sites/${site.id}/pages`)
        .select(['title', 'id'])
        .get();
}

export async function getSitePageWebPartsAsync(site: Site, page: any): Promise<PageCollection> {
    // Ensure client isn't undefined
    if (!_appClient) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    return _appClient?.api(`/sites/${site.id}/pages/${page.id}/microsoft.graph.sitePage/webParts`)
        //.select(['id'])
        .get();
}

export async function getWebpartsOnSites(sites: Site[]): Promise<void> {
    // Ensure client isn't undefined
    if (!_appClient) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    for (const site of sites) {
        console.log(`Site: ${site.webUrl ?? 'NO NAME'}`);
        console.log(`ID: ${site.id}`);

        try {
            const sitePagePage = await getSitePagesAsync(site);
            const pages: any[] = sitePagePage.value;

            for (const page of pages) {
                console.log(` Page: ${page.title ?? 'NO NAME'}`);
                console.log(` ID: ${page.id}`);

                const sitePageWebPart = await getSitePageWebPartsAsync(site, page);
                const webparts: any[] = sitePageWebPart.value;

                try {
                    for (const webpart of webparts) {
                        if (webpart['@odata.type'] == "#microsoft.graph.textWebPart") {
                            continue;
                        }
                        console.log(`  Webpart:`);
                        console.log(`   ID: ${webpart.id}`);
                        console.log(`   webPartType: ${webpart.webPartType}`);
                    }
                }
                catch (err) {
                    console.log(`Error getting webparts: ${err}`);
                }
                
            }

        }
        catch (err) {
            console.log(`Error getting pages: ${err}`);
        }

      } 
      return Promise.resolve();

}