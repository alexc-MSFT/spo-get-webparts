import * as fs from 'fs';

// Read the appsettings.json file
const appSettingsPath = './appsettings.json';
const appSettingsContent = fs.readFileSync(appSettingsPath, 'utf8');
const appSettings = JSON.parse(appSettingsContent);

// Access configuration settings
const clientId = appSettings.clientId;
const clientSecret = appSettings.clientSecret;
const tenantId = appSettings.tenantId;
const webparts = appSettings.webparts;
const sitesInputPath = appSettings.sitesInputPath;
const siteToExcludePath = appSettings.siteToExcludePath;
const allSites : boolean = appSettings.allSites;

console.log('Configuration loaded:', appSettings);

let sites: string[] = [];
let sitesExclude: string[] = [];

const sitesToSearch : string[] = sites;
const sitesToExclude : string[] = sitesExclude;

const settings: AppSettings = {
    clientId: clientId,
    clientSecret: clientSecret,
    tenantId: tenantId,
    webparts: webparts,
    sitesInputPath: sitesInputPath,
    siteToExcludePath: siteToExcludePath,
    sitesToSearch: sitesToSearch,
    sitesToExclude: sitesToExclude,
    allSites: allSites
  };
  
  export interface AppSettings {
    clientId: string;
    clientSecret: string;
    tenantId: string;
    webparts: string[];
    sitesInputPath: string;
    siteToExcludePath: string;
    sitesToSearch : string[];
    sitesToExclude : string[];
    allSites: boolean;
  }
  
  export default settings;