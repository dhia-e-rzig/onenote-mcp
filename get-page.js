#!/usr/bin/env node

import { Client } from '@microsoft/microsoft-graph-client';
import { loadToken, isTokenExpired } from './lib/token-store.js';

// Get the page title from command line
const pageTitle = process.argv[2];
if (!pageTitle) {
  console.error('Please provide a page title as argument. Example: node get-page.js "Questions"');
  process.exit(1);
}

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
    
    // Get all pages
    console.log(`Searching for page with title: "${pageTitle}"...`);
    const pagesResponse = await client.api('/me/onenote/pages').get();
    
    if (!pagesResponse.value || pagesResponse.value.length === 0) {
      console.error('No pages found');
      return;
    }
    
    // Find the requested page
    const page = pagesResponse.value.find(p => 
      p.title && p.title.toLowerCase().includes(pageTitle.toLowerCase())
    );
    
    if (!page) {
      console.error(`No page found with title containing "${pageTitle}"`);
      console.log('Available pages:');
      pagesResponse.value.forEach(p => console.log(`- ${p.title}`));
      return;
    }
    
    console.log(`Found page: "${page.title}" (ID: ${page.id})`);
    
    // Fetch the content
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${page.id}/content`;
    console.log(`Fetching content...`);
    
    const response = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }
    
    const content = await response.text();
    console.log(`Content received! Length: ${content.length} characters`);
    
    // Extract text content
    let plainText = content
      .replace(/<[^>]*>?/gm, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    
    console.log('\n--- PAGE CONTENT ---\n');
    console.log(plainText);
    console.log('\n--- END OF CONTENT ---\n');
    
  } catch (error) {
    console.error('Error:', error.message);
  }
}

getPageContent(); 