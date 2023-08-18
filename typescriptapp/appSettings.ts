import * as fs from 'fs';

// Read the appsettings.json file
const appSettingsPath = './appsettings.json';
const appSettingsContent = fs.readFileSync(appSettingsPath, 'utf8');
const appSettings = JSON.parse(appSettingsContent);

// Access configuration settings
const clientId = appSettings.clientId;
const clientSecret = appSettings.clientSecret;
const tenantId = appSettings.tenantId;
const webparts = appSettings.siteId;

console.log('Configuration loaded:', appSettings);



const settings: AppSettings = {
    clientId: clientId,
    clientSecret: clientSecret,
    tenantId: tenantId,
    webparts: webparts
  };
  
  export interface AppSettings {
    clientId: string;
    clientSecret: string;
    tenantId: string;
    webparts: string[];
  }
  
  export default settings;