#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { InteractiveBrowserCredential } from '@azure/identity';
import fetch from 'node-fetch';

// Import security modules
import { loadToken, saveToken, isTokenExpired, isValidTokenFormat } from './lib/token-store.js';
import { validateSearchTerm, sanitizeHtmlContent, createSafeErrorMessage } from './lib/validation.js';
import { rateLimiter } from './lib/rate-limiter.js';

// Load environment variables
dotenv.config();

// Client ID - Should be set via environment variable for production
// Falls back to Graph Explorer ID for development only
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e';

// Warn if using default client ID
if (!process.env.AZURE_CLIENT_ID) {
  console.error('WARNING: Using default Microsoft Graph Explorer client ID.');
  console.error('For production, register your own Azure AD application and set AZURE_CLIENT_ID environment variable.');
}

// OAuth scopes - use read-only by default, write only when needed
const READ_SCOPES = ['Notes.Read.All', 'User.Read'];
const WRITE_SCOPES = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read'];

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
let currentScopes = READ_SCOPES;

// Credential instance for token refresh
let credential = null;

let graphClient = null;

/**
 * Validate token by making a simple API call
 */
async function validateToken() {
  if (!accessToken) return false;
  if (!isValidTokenFormat(accessToken)) return false;
  
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

/**
 * Authenticate using interactive browser flow
 * @param {string[]} scopes - OAuth scopes to request
 */
async function authenticateInteractive(scopes = READ_SCOPES) {
  console.error('Starting interactive browser authentication...');
  
  credential = new InteractiveBrowserCredential({
    clientId: clientId,
    tenantId: 'consumers',  // Use 'consumers' for personal Microsoft accounts, 'common' for both
    redirectUri: 'http://localhost:8400'
  });
  
  try {
    const tokenResponse = await credential.getToken(scopes);
    accessToken = tokenResponse.token;
    tokenExpiresAt = tokenResponse.expiresOnTimestamp 
      ? new Date(tokenResponse.expiresOnTimestamp) 
      : new Date(Date.now() + 3600 * 1000);
    currentScopes = scopes;
    
    await saveToken(accessToken, tokenExpiresAt);
    console.error('Authentication successful!');
    return true;
  } catch (error) {
    console.error('Authentication failed');
    throw new Error('Authentication failed. Please try again.');
  }
}

/**
 * Ensure Graph client is created with valid token
 * @param {boolean} requireWrite - Whether write permissions are needed
 */
async function ensureGraphClient(requireWrite = false) {
  // Load token from secure store if not already loaded
  if (!accessToken) {
    const stored = await loadToken();
    accessToken = stored.token;
    tokenExpiresAt = stored.expiresAt;
  }
  
  // Determine required scopes
  const requiredScopes = requireWrite ? WRITE_SCOPES : READ_SCOPES;
  
  // Check if we need to authenticate
  let needsAuth = false;
  
  if (!accessToken || !isValidTokenFormat(accessToken)) {
    console.error('No valid token found, authentication required');
    needsAuth = true;
  } else if (isTokenExpired(tokenExpiresAt)) {
    // Token expired - validate it first (Microsoft tokens sometimes work past expiry)
    console.error('Token appears expired, validating...');
    const isValid = await validateToken();
    if (!isValid) {
      console.error('Token validation failed, re-authentication required');
      needsAuth = true;
    } else {
      console.error('Token still valid despite expiry time');
    }
  }
  
  // Only authenticate if we really need to (no token or token is invalid)
  if (needsAuth) {
    await authenticateInteractive(requiredScopes);
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
  async () => {
    try {
      await ensureGraphClient(false);
      const response = await rateLimiter.execute(() => 
        graphClient.api("/me/onenote/notebooks").get()
      );
      return {
        content: [{ type: "text", text: JSON.stringify(response.value) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('List notebooks', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for getting notebook details
server.tool(
  "getNotebook",
  "Get details of a specific notebook",
  async () => {
    try {
      await ensureGraphClient(false);
      const response = await rateLimiter.execute(() => 
        graphClient.api("/me/onenote/notebooks").get()
      );
      
      if (!response.value || response.value.length === 0) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'No notebooks found' }) }]
        };
      }
      
      return { 
        content: [{ type: "text", text: JSON.stringify(response.value[0]) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Get notebook', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for listing sections in a notebook
server.tool(
  "listSections",
  "List all sections in a notebook",
  async () => {
    try {
      await ensureGraphClient(false);
      const response = await rateLimiter.execute(() => 
        graphClient.api("/me/onenote/sections").get()
      );
      return { 
        content: [{ type: "text", text: JSON.stringify(response.value) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('List sections', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for listing pages in a section
server.tool(
  "listPages",
  "List all pages in a section",
  async () => {
    try {
      await ensureGraphClient(false);
      const sectionsResponse = await rateLimiter.execute(() => 
        graphClient.api("/me/onenote/sections").get()
      );
      
      if (!sectionsResponse.value || sectionsResponse.value.length === 0) {
        return { 
          content: [{ type: "text", text: "[]" }]
        };
      }
      
      const sectionId = sectionsResponse.value[0].id;
      const response = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get()
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify(response.value) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('List pages', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for getting the content of a page
server.tool(
  "getPage",
  "Get the content of a page",
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      // Validate input
      const inputValue = params.random_string || '';
      const validation = validateSearchTerm(inputValue);
      if (!validation.valid) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: validation.error }) }]
        };
      }
      
      const pagesResponse = await rateLimiter.execute(() => 
        graphClient.api('/me/onenote/pages').get()
      );
      
      if (!pagesResponse.value || pagesResponse.value.length === 0) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'No pages found' }) }]
        };
      }
      
      let targetPage;
      const searchValue = validation.value;
      
      if (searchValue.length > 0) {
        // Look for exact ID match first
        targetPage = pagesResponse.value.find(p => p.id === searchValue);
        
        // If no exact match, try matching by title
        if (!targetPage) {
          const searchLower = searchValue.toLowerCase();
          targetPage = pagesResponse.value.find(p => 
            p.title && p.title.toLowerCase().includes(searchLower)
          );
        }
      } else {
        targetPage = pagesResponse.value[0];
      }
      
      if (!targetPage) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'Page not found' }) }]
        };
      }
      
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage.id}/content`;
      
      const content = await rateLimiter.execute(async () => {
        const res = await fetch(url, {
          headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        if (!res.ok) {
          throw new Error(`HTTP ${res.status}`);
        }
        return res.text();
      });
      
      return {
        content: [{ type: "text", text: content }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Get page', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for creating a new page in a section
server.tool(
  "createPage",
  "Create a new page in a section",
  async () => {
    try {
      // Require write permissions for this operation
      await ensureGraphClient(true);
      
      const sectionsResponse = await rateLimiter.execute(() => 
        graphClient.api("/me/onenote/sections").get()
      );
      
      if (!sectionsResponse.value || sectionsResponse.value.length === 0) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'No sections found' }) }]
        };
      }
      
      const sectionId = sectionsResponse.value[0].id;
      
      // Sanitize HTML content
      const simpleHtml = sanitizeHtmlContent(`
        <!DOCTYPE html>
        <html>
          <head>
            <title>New Page</title>
          </head>
          <body>
            <p>This is a new page created via the Microsoft Graph API</p>
          </body>
        </html>
      `);
      
      const response = await rateLimiter.execute(() => 
        graphClient
          .api(`/me/onenote/sections/${sectionId}/pages`)
          .header("Content-Type", "application/xhtml+xml")
          .post(simpleHtml)
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify({ success: true, id: response.id }) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Create page', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for searching pages
server.tool(
  "searchPages",
  "Search for pages across notebooks",
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      // Validate search input
      const validation = validateSearchTerm(params.random_string);
      if (!validation.valid) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: validation.error }) }]
        };
      }
      
      const response = await rateLimiter.execute(() => 
        graphClient.api("/me/onenote/pages").get()
      );
      
      if (!response.value) {
        return { 
          content: [{ type: "text", text: "[]" }]
        };
      }
      
      // Filter by search term if provided
      if (validation.value.length > 0) {
        const searchTerm = validation.value.toLowerCase();
        const filteredPages = response.value.filter(page => 
          page.title && page.title.toLowerCase().includes(searchTerm)
        );
        
        return { 
          content: [{ type: "text", text: JSON.stringify(filteredPages) }]
        };
      }
      
      return { 
        content: [{ type: "text", text: JSON.stringify(response.value) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Search pages', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
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
    console.error('Server failed to start');
    process.exit(1);
  }
}

main(); 