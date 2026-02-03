import { Client } from '@microsoft/microsoft-graph-client';
import { loadToken, isTokenExpired } from './lib/token-store.js';

async function listSections() {
  try {
    // Load token from secure store
    const { token: accessToken, expiresAt } = await loadToken();
    
    if (!accessToken) {
      console.error('No access token found. Please run: node authenticate.js');
      return;
    }
    
    if (isTokenExpired(expiresAt)) {
      console.error('Token expired. Please run: node authenticate.js');
      return;
    }

    // Create Microsoft Graph client
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // First, let's get all notebooks
    console.log('Fetching notebooks...');
    const notebooksResponse = await client.api('/me/onenote/notebooks').get();
    
    if (notebooksResponse.value.length === 0) {
      console.log('No notebooks found.');
      return;
    }

    const notebook = notebooksResponse.value[0];
    console.log(`Using notebook: "${notebook.displayName}"`);

    // Get sections in the selected notebook
    console.log(`Fetching sections in "${notebook.displayName}" notebook...`);
    const sectionsResponse = await client.api(`/me/onenote/notebooks/${notebook.id}/sections`).get();
    
    console.log(`\nSections in ${notebook.displayName} Notebook:`);
    console.log('============================');
    
    if (sectionsResponse.value.length === 0) {
      console.log('No sections found.');
    } else {
      sectionsResponse.value.forEach((section, index) => {
        console.log(`${index + 1}. ${section.displayName}`);
      });
    }

  } catch (error) {
    console.error('Error listing sections:', error.message);
  }
}

listSections(); 