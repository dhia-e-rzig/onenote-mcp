import { Client } from '@microsoft/microsoft-graph-client';
import { loadToken, isTokenExpired } from './lib/token-store.js';

async function createPage() {
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

    // Create a new page
    console.log(`Creating a new page in "${section.displayName}" section...`);
    
    const now = new Date();
    const formattedDate = now.toISOString().split('T')[0];
    const formattedTime = now.toLocaleTimeString();
    
    const simpleHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>Created via MCP on ${formattedDate}</title>
        </head>
        <body>
          <h1>Created via MCP on ${formattedDate}</h1>
          <p>This page was created via the Microsoft Graph API at ${formattedTime}.</p>
          <p>This demonstrates that the OneNote MCP integration is working correctly!</p>
          <ul>
            <li>The authentication flow is working</li>
            <li>We can create new pages</li>
            <li>We can access existing notebooks</li>
          </ul>
        </body>
      </html>
    `;
    
    const response = await client
      .api(`/me/onenote/sections/${section.id}/pages`)
      .header("Content-Type", "application/xhtml+xml")
      .post(simpleHtml);
    
    console.log(`\nNew page created successfully:`);
    console.log(`Title: ${response.title}`);
    console.log(`Created: ${new Date(response.createdDateTime).toLocaleString()}`);
    console.log(`Link: ${response.links.oneNoteWebUrl.href}`);

  } catch (error) {
    console.error('Error creating page:', error.message);
  }
}

createPage(); 