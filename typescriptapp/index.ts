import * as readline from 'readline-sync';
import { Site } from '@microsoft/microsoft-graph-types';
import * as fs from 'fs';
import * as graphHelper from './graphHelper';

import settings, { AppSettings } from './appSettings';

async function main() {

  if (settings.allSites) {
    console.log('All sites will be checked');
    console.log('Checking if there are any sites to exclude');
    try {
      // Read the entire file as a string
      const fileContents = fs.readFileSync(settings.siteToExcludePath, 'utf-8');

      // Split the file contents into lines using the newline character as a delimiter
      settings.sitesToExclude = fileContents.split('\n').map(line => line.trim());

      console.log(`Sites To Exclude loaded from ${settings.siteToExcludePath}`);

    } catch (error) {
      console.error('Error reading the sites exclude file:', error);
      settings.sitesToExclude = [];
    }
  } else {

    // Read the entire file as a string
    const fileContents = fs.readFileSync(settings.sitesInputPath, 'utf-8');

    // Split the file contents into lines using the newline character as a delimiter
    settings.sitesToSearch = fileContents.split('\n').map(line => line.trim());

    console.log(`Sites Input loaded from ${settings.sitesInputPath}`);
  }

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  const choices = [
    'Display access token',
    'List all sites',
    'List webparts on sites'
  ];

  // Generate a timestamp for the file name
  const timestamp = new Date().toISOString().replace(/[-:T.]/g, '');

  // Specify the file path with the timestamp
  const fileName = `WebPartsInSite_${timestamp}.csv`;
  const filePath = fileName;

  //await listAllSitesAsync();
  //await listWebPartsOnSitesAsync(settings, filePath);

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        // Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List users
        await listAllSitesAsync();
        break;
      case 2:
        // Run any Graph code
        await listWebPartsOnSitesAsync(settings, filePath);
        break;
      default:
        console.log('Invalid choice! Please try again.');
    }
  }
}

function initializeGraph(settings: AppSettings) {
  graphHelper.initializeGraphForAppOnlyAuth(settings);
}

async function displayAccessTokenAsync() {
  try {
    const userToken = await graphHelper.getAppOnlyTokenAsync();
    console.log(`App-only token: ${userToken}`);
    console.log(settings);
  } catch (err) {
    console.log(`Error getting app-only access token: ${err}`);
  }
}

async function listAllSitesAsync() {
  try {

    let sites: Site[] = [];

    if (settings.allSites) {
      sites = await graphHelper.getAllSitesAsync();
    } else {

      for (const siteUrl of settings.sitesToSearch) {

        try {
          const site = await graphHelper.getSiteAsync(siteUrl);
          sites.push(site);
        }
        catch (err) {
          console.log(`Error getting site: ${siteUrl}`);
          console.log(`${err}`);
        }
      }

    }



    console.log(`Sites Found: ${sites.length}`);

  } catch (err) {
    console.log(`Error getting sites: ${err}`);
  }
}

async function listSitesAsync() {
  try {
    const sitePage = await graphHelper.getSitesAsync();
    const sites: Site[] = sitePage.value;

    // Output each user's details
    for (const site of sites) {
      console.log(`Site: ${site.webUrl ?? 'NO NAME'}`);
      console.log(`  ID: ${site.id}`);
    }

    // If @odata.nextLink is not undefined, there are more users
    // available on the server
    const moreAvailable = sitePage['@odata.nextLink'] != undefined;
    console.log(`\nMore sites available? ${moreAvailable}`);
  } catch (err) {
    console.log(`Error getting sites: ${err}`);
  }
}

async function listWebPartsOnSitesAsync(settings: AppSettings, filePath: string) {
  try {
    let sites: Site[] = [];

    if (settings.allSites) {
      sites = await graphHelper.getAllSitesAsync();
    } else {

      for (const siteUrl of settings.sitesToSearch) {
        try {
          const site = await graphHelper.getSiteAsync(siteUrl);
          sites.push(site);
        }
        catch (err) {
          console.log(`Error getting site: ${siteUrl}`);
          console.log(`${err}`);
        }
      }

    }

    console.log(`Sites Found: ${sites.length}`);

    await graphHelper.getWebpartsOnSites(sites, settings.webparts, filePath);


  } catch (err) {
    console.log(`Error getting sites: ${err}`);
  }
}

main();

