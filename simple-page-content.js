#!/usr/bin/env node

import { Client } from '@microsoft/microsoft-graph-client';
import { loadToken, isTokenExpired } from './lib/token-store.js';

// Main function
async function getPageContent() {
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
    
    // Initialize Graph client
    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
    
    // List pages
    console.log("Fetching pages...");
    const pages = await client.api('/me/onenote/pages').get();
    
    if (!pages || !pages.value || pages.value.length === 0) {
      console.log("No pages found");
      return;
    }
    
    // Choose the first page
    const page = pages.value[0];
    console.log(`Using page: "${page.title}" (ID: ${page.id})`);
    
    // Try to get the content
    console.log("Fetching page content...");
    
    try {
      // Create direct HTTP request to the content endpoint
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${page.id}/content`;
      
      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`
        }
      });
      
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      
      const contentType = response.headers.get('content-type');
      console.log(`Content type: ${contentType}`);
      
      const content = await response.text();
      console.log(`Content received! Length: ${content.length} characters`);
      console.log(`Content preview (first 100 chars): ${content.substring(0, 100).replace(/\n/g, ' ')}...`);
      
      // Don't save content to file - just confirm it worked
      console.log("Content retrieval successful! Privacy preserved - not saving to disk.");
    } catch (error) {
      console.error("Error fetching content:", error.message);
    }
    
  } catch (error) {
    console.error("Error:", error.message);
  }
}

getPageContent(); 