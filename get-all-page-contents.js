#!/usr/bin/env node

import { Client } from '@microsoft/microsoft-graph-client';
import { JSDOM } from 'jsdom';
import { loadToken, isTokenExpired } from './lib/token-store.js';

// Function to extract text content from HTML (limited output for security)
function extractTextContent(html) {
  try {
    const dom = new JSDOM(html);
    const document = dom.window.document;
    
    const bodyText = document.body.textContent.trim();
    
    // Limit output to first 300 chars for summary
    const summary = bodyText.substring(0, 300).replace(/\s+/g, ' ');
    
    return summary.length < bodyText.length 
      ? `${summary}...` 
      : summary;
  } catch (error) {
    return 'Could not extract text content';
  }
}

// Main function
async function getAllPageContents() {
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
    
    console.log(`Found ${pages.value.length} pages. Fetching content for each...`);
    
    // Process each page
    for (const page of pages.value) {
      console.log(`\n===== ${page.title} =====`);
      
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
        const textSummary = extractTextContent(content);
        
        console.log(`Last modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}`);
        console.log(`Content summary: ${textSummary}`);
      } catch (error) {
        console.error(`Error processing ${page.title}`);
      }
    }
    
  } catch (error) {
    console.error("Error:", error.message);
  }
}

getAllPageContents(); 