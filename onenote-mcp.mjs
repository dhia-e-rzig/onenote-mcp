#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import { InteractiveBrowserCredential } from '@azure/identity';
import fetch from 'node-fetch';

// Load environment variables
dotenv.config();

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Create the MCP server
const server = new McpServer(
  { 
    name: "onenote",
    version: "1.0.0",
    description: "OneNote MCP Server" 
  },
  {
    capabilities: {
      tools: {
        listChanged: true
      }
    }
  }
);

// Token storage with expiry tracking
let accessToken = null;
let tokenExpiresAt = null;

// Client ID for Microsoft Graph API access
const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Explorer client ID
const scopes = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read'];

// Credential instance for token refresh
let credential = null;

let graphClient = null;

// Load token from file
function loadTokenFromFile() {
  try {
    if (fs.existsSync(tokenFilePath)) {
      const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
      try {
        const parsedToken = JSON.parse(tokenData);
        accessToken = parsedToken.token;
        tokenExpiresAt = parsedToken.expiresAt ? new Date(parsedToken.expiresAt) : null;
      } catch (parseError) {
        accessToken = tokenData;
        tokenExpiresAt = null;
      }
    }
  } catch (error) {
    console.error('Error reading access token file:', error.message);
  }
}

// Save token to file
function saveTokenToFile(token, expiresAt) {
  const tokenData = JSON.stringify({ 
    token: token, 
    expiresAt: expiresAt ? expiresAt.toISOString() : null 
  });
  fs.writeFileSync(tokenFilePath, tokenData);
}

// Check if token is expired or about to expire (within 5 minutes)
function isTokenExpired() {
  if (!tokenExpiresAt) return true;
  const bufferTime = 5 * 60 * 1000; // 5 minutes
  return new Date().getTime() > (tokenExpiresAt.getTime() - bufferTime);
}

// Validate token by making a simple API call
async function validateToken() {
  if (!accessToken) return false;
  
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });
    return response.ok;
  } catch (error) {
    return false;
  }
}

// Authenticate using interactive browser flow
async function authenticateInteractive() {
  console.error('Starting interactive browser authentication...');
  console.error('A browser window will open for you to sign in with your Microsoft account...');
  
  credential = new InteractiveBrowserCredential({
    clientId: clientId,
    redirectUri: 'http://localhost:8400'
  });
  
  try {
    const tokenResponse = await credential.getToken(scopes);
    accessToken = tokenResponse.token;
    tokenExpiresAt = tokenResponse.expiresOnTimestamp 
      ? new Date(tokenResponse.expiresOnTimestamp) 
      : new Date(Date.now() + 3600 * 1000); // Default 1 hour
    
    saveTokenToFile(accessToken, tokenExpiresAt);
    console.error('Authentication successful!');
    return true;
  } catch (error) {
    console.error('Authentication error:', error.message);
    throw new Error(`Authentication failed: ${error.message}`);
  }
}

// Refresh token using existing credential
async function refreshToken() {
  if (!credential) {
    return false;
  }
  
  try {
    console.error('Attempting to refresh token...');
    const tokenResponse = await credential.getToken(scopes);
    accessToken = tokenResponse.token;
    tokenExpiresAt = tokenResponse.expiresOnTimestamp 
      ? new Date(tokenResponse.expiresOnTimestamp) 
      : new Date(Date.now() + 3600 * 1000);
    
    saveTokenToFile(accessToken, tokenExpiresAt);
    console.error('Token refreshed successfully!');
    return true;
  } catch (error) {
    console.error('Token refresh failed:', error.message);
    return false;
  }
}

// Function to ensure Graph client is created with valid token
async function ensureGraphClient() {
  // Load token from file if not already loaded
  if (!accessToken) {
    loadTokenFromFile();
  }
  
  // Also check environment variable
  if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
    accessToken = process.env.GRAPH_ACCESS_TOKEN;
  }
  
  // Check if we need to authenticate or refresh
  let needsAuth = false;
  
  if (!accessToken) {
    needsAuth = true;
  } else if (isTokenExpired()) {
    // Try to refresh first
    const refreshed = await refreshToken();
    if (!refreshed) {
      const isValid = await validateToken();
      if (!isValid) {
        needsAuth = true;
      }
    }
  } else {
    // Token exists and not expired, but validate it
    const isValid = await validateToken();
    if (!isValid) {
      const refreshed = await refreshToken();
      if (!refreshed) {
        needsAuth = true;
      }
    }
  }
  
  // If we need to authenticate, do it now (blocking)
  if (needsAuth) {
    await authenticateInteractive();
  }
  
  // Create or recreate the Graph client with current token
  graphClient = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
  
  return graphClient;
}

// Tool for listing all notebooks
server.tool(
  "listNotebooks",
  "List all OneNote notebooks",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/notebooks").get();
      // Return content as an array of text items
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing notebooks:", error);
      throw new Error(`Failed to list notebooks: ${error.message}`);
    }
  }
);

// Tool for getting notebook details
server.tool(
  "getNotebook",
  "Get details of a specific notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/notebooks`).get();
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value[0])
          }
        ]
      };
    } catch (error) {
      console.error("Error getting notebook:", error);
      throw new Error(`Failed to get notebook: ${error.message}`);
    }
  }
);

// Tool for listing sections in a notebook
server.tool(
  "listSections",
  "List all sections in a notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/sections`).get();
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing sections:", error);
      throw new Error(`Failed to list sections: ${error.message}`);
    }
  }
);

// Tool for listing pages in a section
server.tool(
  "listPages",
  "List all pages in a section",
  async (params) => {
    try {
      await ensureGraphClient();
      // Get sections first
      const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();
      
      if (sectionsResponse.value.length === 0) {
        return { 
          content: [
            {
              type: "text",
              text: "[]"
            }
          ]
        };
      }
      
      // Use the first section
      const sectionId = sectionsResponse.value[0].id;
      const response = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get();
      
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing pages:", error);
      throw new Error(`Failed to list pages: ${error.message}`);
    }
  }
);

// Tool for getting the content of a page
server.tool(
  "getPage",
  "Get the content of a page",
  async (params) => {
    try {
      await ensureGraphClient();
      
      // First, list all pages to find the one we want
      const pagesResponse = await graphClient.api('/me/onenote/pages').get();
      
      let targetPage;
      
      // If a page ID is provided, use it to find the page
      if (params.random_string && params.random_string.length > 0) {
        const pageId = params.random_string;
        
        // Look for exact match first
        targetPage = pagesResponse.value.find(p => p.id === pageId);
        
        // If no exact match, try matching by title
        if (!targetPage) {
          targetPage = pagesResponse.value.find(p => 
            p.title && p.title.toLowerCase().includes(params.random_string.toLowerCase())
          );
        }
        
        // If still no match, try partial ID match
        if (!targetPage) {
          targetPage = pagesResponse.value.find(p => 
            p.id.includes(pageId) || pageId.includes(p.id)
          );
        }
      } else {
        // If no ID provided, use the first page
        targetPage = pagesResponse.value[0];
      }
      
      if (!targetPage) {
        throw new Error("Page not found");
      }
      
      try {
        const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage.id}/content`;
        
        // Make direct HTTP request with fetch
        const response = await fetch(url, {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        });
        
        if (!response.ok) {
          throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
        }
        
        const content = await response.text();
        
        // Return the raw HTML content
        return {
          content: [
            {
              type: "text",
              text: content
            }
          ]
        };
      } catch (error) {
        return {
          content: [
            {
              type: "text",
              text: `Error retrieving page content: ${error.message}`
            }
          ]
        };
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error in getPage: ${error.message}`
          }
        ]
      };
    }
  }
);

// Tool for creating a new page in a section
server.tool(
  "createPage",
  "Create a new page in a section",
  async (params) => {
    try {
      await ensureGraphClient();
      // Get sections first
      const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();
      
      if (sectionsResponse.value.length === 0) {
        throw new Error("No sections found");
      }
      
      // Use the first section
      const sectionId = sectionsResponse.value[0].id;
      
      // Create simple HTML content
      const simpleHtml = `
        <!DOCTYPE html>
        <html>
          <head>
            <title>New Page</title>
          </head>
          <body>
            <p>This is a new page created via the Microsoft Graph API</p>
          </body>
        </html>
      `;
      
      const response = await graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(simpleHtml);
      
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response)
          }
        ]
      };
    } catch (error) {
      console.error("Error creating page:", error);
      throw new Error(`Failed to create page: ${error.message}`);
    }
  }
);

// Tool for searching pages
server.tool(
  "searchPages",
  "Search for pages across notebooks",
  async (params) => {
    try {
      await ensureGraphClient();
      
      // Get all pages
      const response = await graphClient.api(`/me/onenote/pages`).get();
      
      // If search string is provided, filter the results
      if (params.random_string && params.random_string.length > 0) {
        const searchTerm = params.random_string.toLowerCase();
        const filteredPages = response.value.filter(page => {
          // Search in title
          if (page.title && page.title.toLowerCase().includes(searchTerm)) {
            return true;
          }
          return false;
        });
        
        return { 
          content: [
            {
              type: "text",
              text: JSON.stringify(filteredPages)
            }
          ]
        };
      } else {
        // Return all pages if no search term
        return { 
          content: [
            {
              type: "text",
              text: JSON.stringify(response.value)
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error searching pages:", error);
      throw new Error(`Failed to search pages: ${error.message}`);
    }
  }
);

// Connect to stdio and start server
async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('OneNote MCP Server started successfully.');
    console.error('Authentication will be triggered automatically when needed.');
    
    process.on('SIGINT', () => {
      process.exit(0);
    });
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

main(); 