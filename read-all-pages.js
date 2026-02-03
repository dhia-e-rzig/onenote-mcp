#!/usr/bin/env node

import { Client } from '@microsoft/microsoft-graph-client';
import { JSDOM } from 'jsdom';
import { loadToken, isTokenExpired } from './lib/token-store.js';

// Extract readable text from HTML
function extractReadableText(html) {
  try {
    const dom = new JSDOM(html);
    const document = dom.window.document;
    
    // Remove scripts
    const scripts = document.querySelectorAll('script');
    scripts.forEach(script => script.remove());
    
    let text = '';
    
    // Process headings
    document.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach(heading => {
      text += `\n${heading.textContent.trim()}\n${'-'.repeat(heading.textContent.length)}\n`;
    });
    
    // Process paragraphs
    document.querySelectorAll('p').forEach(paragraph => {
      const content = paragraph.textContent.trim();
      if (content) {
        text += `${content}\n\n`;
      }
    });
    
    // Process lists
    document.querySelectorAll('ul, ol').forEach(list => {
      text += '\n';
      list.querySelectorAll('li').forEach((item, index) => {
        const content = item.textContent.trim();
        if (content) {
          text += `${index + 1}. ${content}\n`;
        }
      });
      text += '\n';
    });
    
    // Process tables
    document.querySelectorAll('table').forEach(table => {
      text += '\nTable content:\n';
      table.querySelectorAll('tr').forEach(row => {
        const cells = Array.from(row.querySelectorAll('td, th'))
          .map(cell => cell.textContent.trim())
          .join(' | ');
        text += `${cells}\n`;
      });
      text += '\n';
    });
    
    // Fallback
    if (!text.trim()) {
      text = document.body.textContent.trim().replace(/\s+/g, ' ');
    }
    
    return text;
  } catch (error) {
    return 'Could not extract readable text from HTML content.';
  }
}

// Main function
async function readAllPages() {
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
    
    console.log(`Found ${pages.value.length} pages. Reading full content for each...\n`);
    
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
        
        const htmlContent = await response.text();
        const readableText = extractReadableText(htmlContent);
        
        console.log(readableText);
        console.log("\n");
      } catch (error) {
        console.error(`Error processing ${page.title}`);
      }
    }
    
    console.log("\nAll pages have been read. You can now ask questions about their content.");
    
  } catch (error) {
    console.error("Error:", error.message);
  }
}

readAllPages(); 