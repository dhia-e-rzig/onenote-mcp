#!/usr/bin/env node

import { Client } from '@microsoft/microsoft-graph-client';
import { loadToken, isTokenExpired } from './lib/token-store.js';

// Main function
async function getAllPagesFullContent() {
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
    console.log("Fetching all pages...");
    const pages = await client.api('/me/onenote/pages').get();
    
    if (!pages || !pages.value || pages.value.length === 0) {
      console.log("No pages found");
      return;
    }
    
    console.log(`Found ${pages.value.length} pages. Fetching full content for each...\n`);
    
    // Process each page
    for (const page of pages.value) {
      console.log(`\n==================================================================`);
      console.log(`PAGE: ${page.title}`);
      console.log(`Last modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}`);
      console.log(`==================================================================\n`);
      
      try {
        const url = page.contentUrl;
        
        const response = await fetch(url, {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        });
        
        if (!response.ok) {
          console.error(`Error fetching ${page.title}: ${response.status}`);
          continue;
        }
        
        const content = await response.text();
        
        console.log("FULL HTML CONTENT:");
        console.log(content);
        console.log("\n");
      } catch (error) {
        console.error(`Error processing ${page.title}`);
      }
    }
    
  } catch (error) {
    console.error("Error:", error.message);
  }
}

getAllPagesFullContent(); 