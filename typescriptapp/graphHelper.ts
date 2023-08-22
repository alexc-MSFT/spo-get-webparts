import 'isomorphic-fetch';
import { ClientSecretCredential } from '@azure/identity';
import { Client, GraphRequestOptions, PageCollection, PageIterator, PageIteratorCallback } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from
    '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

import { AppSettings } from './appSettings.ts';
import { Site } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';

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
        // console.log(site.webUrl);
        if (site.webUrl?.includes('my.sharepoint.com')) {
            // don't add mysite
        } else {

            sites.push(site);
        }

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

export async function getWebpartsOnSites(sites: Site[], webpartIdsToFind: string[], filePath: string): Promise<void> {
    // Ensure client isn't undefined
    if (!_appClient) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    // Get name
    let splitPath = filePath.split('.');
    let fileName = splitPath[0];
    let fileExtension = splitPath[1];

    // Create a writable stream to append data to the file
    const successStream = fs.createWriteStream(`${fileName}_Success.${fileExtension}`, { flags: 'a' });
    const errorStream = fs.createWriteStream(`${fileName}_Error.${fileExtension}`, { flags: 'a' });

    // ADd handler for stream events
    successStream.on('finish', () => {
        console.log('CSV data appended successfully!');
    });
    errorStream.on('finish', () => {
        console.log('CSV data appended successfully!');
    });

    successStream.on('error', (err) => {
        console.error('Error appending CSV data:', err);
    });

    errorStream.on('error', (err) => {
        console.error('Error appending CSV data:', err);
    });

    // Header row
    const header = ['Timestamp', 'SiteUrl', 'Page', 'Webpart', 'WebpartName', 'Type', 'Notes'];
    const csvRow = header.join(',') + '\n';
    successStream.write(csvRow);
    errorStream.write(csvRow);

    for (const site of sites) {
        // console.log(`Site: ${site.webUrl ?? 'NO NAME'}`);
        // console.log(`ID: ${site.id}`);

        try {

            // get subsites
            let subsites = await getSubsites(site);

            // create a list of the site and it's subsites
            let sitesToSearch = [site, ...subsites];


            for (const siteToSearch of sitesToSearch) {
                // console.log(`Site: ${siteToSearch.webUrl ?? 'NO NAME'}`);

                try {
                    const sitePagePage = await getSitePagesAsync(siteToSearch);
                    const pages: any[] = sitePagePage.value;

                    for (const page of pages) {
                        // console.log(` Page: ${page.title ?? 'NO NAME'}`);
                        // console.log(` ID: ${page.id}`);

                        try {
                            const sitePageWebPart = await getSitePageWebPartsAsync(siteToSearch, page);
                            const webparts: any[] = sitePageWebPart.value;


                            for (const webpart of webparts) {
                                if (webpart['@odata.type'] == "#microsoft.graph.textWebPart") {
                                    continue;
                                }

                                if (webpartIdsToFind.indexOf(webpart.webPartType) === -1) {
                                    continue;
                                }


                                WriteLog(successStream, siteToSearch, page, "", webpart, "Success");

                                // Get site Owners
                                // 
                                
                            }
                        }
                        catch (err) {
                            WriteLog(errorStream, siteToSearch, page, err, "", "Error Getting Webparts");
                            console.log(`Error getting webparts: ${err}`);

                        }

                    }
                }
                catch (err) {
                    WriteLog(errorStream, siteToSearch, "", err, "", "Error Getting Pages");
                    console.log(`Error getting pages: ${err}`);
                }


            }
        }

        catch (err) {
            WriteLog(errorStream, site, "", err, "", "Error Getting SubSites");
            console.log(`Error getting pages: ${err}`);
        }


    }
    // Close the stream when done appending
    errorStream.end();
    successStream.end();

    return Promise.resolve();
}

async function getSubsites(site: Site) {
    // Ensure client isn't undefined
    if (!_appClient) {
        throw new Error('Graph has not been initialized for app-only auth');
    }

    let subsites = [];

    let sites = await _appClient.api(`/sites/${site.id}/sites`)
        .select(['webUrl', 'id'])
        .get();
    for (let subsite of sites.value) {
        subsites.push(subsite);

        let subsubsites: any = await getSubsites(subsite);
        subsites.push(...subsubsites);
    }
    return subsites;
}

function WriteLog(stream: fs.WriteStream, site: Site, page: any, err: any, webpart: any, type: string) {
    const timestamp = new Date().toISOString().replace(/[-:T.]/g, '');
    let webpartFriendlyName = ""
    if (webpart.data !== undefined) {
        webpartFriendlyName = webpart.data.title;
    }
    const row = [timestamp, site.webUrl, page.title, webpart.webPartType, webpartFriendlyName, type, err];
    const csvRow = row.join(',') + '\n';
    stream.write(csvRow);
}