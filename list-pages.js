import { Client } from '@microsoft/microsoft-graph-client';
import { loadToken, isTokenExpired } from './lib/token-store.js';

async function listPages() {
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

    // First, get all notebooks
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
    
    if (sectionsResponse.value.length === 0) {
      console.log('No sections found in this notebook.');
      return;
    }

    const section = sectionsResponse.value[0];
    console.log(`Using section: "${section.displayName}"`);

    // Get pages in the section
    console.log(`Fetching pages in "${section.displayName}" section...`);
    const pagesResponse = await client.api(`/me/onenote/sections/${section.id}/pages`).get();
    
    console.log(`\nPages in ${section.displayName}:`);
    console.log('=====================');
    
    if (pagesResponse.value.length === 0) {
      console.log('No pages found.');
    } else {
      pagesResponse.value.forEach((page, index) => {
        console.log(`${index + 1}. ${page.title} (Created: ${new Date(page.createdDateTime).toLocaleString()})`);
      });
    }

  } catch (error) {
    console.error('Error listing pages:', error.message);
  }
}

listPages(); 