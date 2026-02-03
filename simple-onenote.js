import { Client } from '@microsoft/microsoft-graph-client';
import { loadToken, isTokenExpired } from './lib/token-store.js';

async function listNotebooks() {
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

    // Get notebooks
    console.log('Fetching notebooks...');
    const response = await client.api('/me/onenote/notebooks').get();
    
    console.log('\nYour OneNote Notebooks:');
    console.log('=======================');
    
    if (response.value.length === 0) {
      console.log('No notebooks found.');
    } else {
      response.value.forEach((notebook, index) => {
        console.log(`${index + 1}. ${notebook.displayName}`);
      });
    }

  } catch (error) {
    console.error('Error listing notebooks:', error.message);
  }
}

listNotebooks(); 