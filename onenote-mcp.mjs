#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { PublicClientApplication, CryptoProvider } from '@azure/msal-node';
import http from 'http';
import fetch from 'node-fetch';
import open from 'open';

// Import security modules
import { loadToken, saveToken, loadRefreshToken, saveRefreshToken, saveAccountInfo, isTokenExpired, isValidTokenFormat } from './lib/token-store.js';
import { validateSearchTerm, sanitizeHtmlContent, createSafeErrorMessage } from './lib/validation.js';
import { rateLimiter } from './lib/rate-limiter.js';

// Load environment variables
dotenv.config();

// Client ID - Should be set via environment variable for production
const clientId = process.env.AZURE_CLIENT_ID || '813d941f-92ac-4ac0-94a2-1e89b720e15b';

// Warn if using default client ID
if (!process.env.AZURE_CLIENT_ID) {
  console.error('WARNING: Using default Microsoft Graph Explorer client ID.');
  console.error('For production, register your own Azure AD application and set AZURE_CLIENT_ID environment variable.');
}

// OAuth scopes - use read-only by default, write only when needed
// Include offline_access for refresh tokens
const READ_SCOPES = ['Notes.Read.All', 'User.Read', 'offline_access'];
const WRITE_SCOPES = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read', 'offline_access'];

// MSAL configuration
const msalConfig = {
  auth: {
    clientId: clientId,
    authority: 'https://login.microsoftonline.com/consumers'  // Use 'common' for both personal and work accounts
  }
};

const pca = new PublicClientApplication(msalConfig);
const cryptoProvider = new CryptoProvider();

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
 * Authenticate using interactive browser flow with local server
 * @param {string[]} scopes - OAuth scopes to request
 */
async function authenticateInteractive(scopes = READ_SCOPES) {
  console.error('Starting interactive browser authentication...');
  
  // Generate PKCE codes before starting the flow
  const pkceCodes = await cryptoProvider.generatePkceCodes();
  
  return new Promise((resolve, reject) => {
    const redirectUri = 'http://localhost:8400';
    
    // Create a local server to handle the redirect
    const server = http.createServer(async (req, res) => {
      const url = new URL(req.url, redirectUri);
      
      if (url.pathname === '/' && url.searchParams.has('code')) {
        const code = url.searchParams.get('code');
        
        try {
          // Exchange code for token with PKCE verifier
          const tokenRequest = {
            code: code,
            scopes: scopes,
            redirectUri: redirectUri,
            codeVerifier: pkceCodes.verifier
          };
          
          const response = await pca.acquireTokenByCode(tokenRequest);
          
          accessToken = response.accessToken;
          tokenExpiresAt = response.expiresOn || new Date(Date.now() + 3600 * 1000);
          currentScopes = scopes;
          
          await saveToken(accessToken, tokenExpiresAt);
          
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end('<html><body><h1>Authentication successful!</h1><p>You can close this window.</p><script>window.close();</script></body></html>');
          
          server.close();
          console.error('Authentication successful!');
          resolve(true);
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'text/html' });
          res.end('<html><body><h1>Authentication failed</h1><p>Please try again.</p></body></html>');
          server.close();
          reject(new Error('Failed to exchange code for token'));
        }
      } else if (url.pathname === '/' && url.searchParams.has('error')) {
        const error = url.searchParams.get('error');
        const errorDescription = url.searchParams.get('error_description') || 'Unknown error';
        
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(`<html><body><h1>Authentication failed</h1><p>${errorDescription}</p></body></html>`);
        server.close();
        reject(new Error(errorDescription));
      }
    });
    
    server.listen(8400, async () => {
      console.error('Waiting for authentication on http://localhost:8400...');
      
      // Generate auth URL and open browser with PKCE
      const authCodeUrlParameters = {
        scopes: scopes,
        redirectUri: redirectUri,
        codeChallenge: pkceCodes.challenge,
        codeChallengeMethod: 'S256'
      };
      
      try {
        const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
        console.error('Opening browser for authentication...');
        await open(authUrl);
      } catch (error) {
        server.close();
        reject(new Error('Failed to generate auth URL'));
      }
    });
    
    // Timeout after 2 minutes
    setTimeout(() => {
      server.close();
      reject(new Error('Authentication timed out'));
    }, 120000);
  });
}

/**
 * Refresh the access token using the stored refresh token
 * @returns {Promise<boolean>} - True if refresh succeeded
 */
async function refreshAccessToken() {
  const refreshToken = await loadRefreshToken();
  if (!refreshToken) {
    console.error('No refresh token available');
    return false;
  }
  
  try {
    console.error('Attempting to refresh access token...');
    
    const refreshRequest = {
      refreshToken: refreshToken,
      scopes: READ_SCOPES
    };
    
    const response = await pca.acquireTokenByRefreshToken(refreshRequest);
    
    accessToken = response.accessToken;
    tokenExpiresAt = response.expiresOn || new Date(Date.now() + 3600 * 1000);
    
    // Save new tokens
    await saveToken(accessToken, tokenExpiresAt);
    
    // Save new refresh token if provided (token rotation)
    if (response.refreshToken) {
      await saveRefreshToken(response.refreshToken);
    }
    
    // Update account info
    if (response.account) {
      await saveAccountInfo(response.account);
    }
    
    console.error('Token refreshed successfully, expires:', tokenExpiresAt);
    return true;
  } catch (error) {
    console.error('Failed to refresh token:', error.message);
    return false;
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
    console.error('Loaded token from store, expires:', tokenExpiresAt);
  }
  
  // Check if we have a valid token format
  if (!accessToken || !isValidTokenFormat(accessToken)) {
    // Try to refresh using stored refresh token
    const refreshed = await refreshAccessToken();
    if (!refreshed) {
      throw new Error('No valid token found. Please run "node authenticate.js" first to sign in.');
    }
  }
  
  // If token appears expired, try to refresh it first
  if (isTokenExpired(tokenExpiresAt)) {
    console.error('Token appears expired, attempting refresh...');
    const refreshed = await refreshAccessToken();
    
    if (!refreshed) {
      // Refresh failed, validate current token with API call
      console.error('Refresh failed, validating current token with API...');
      const isValid = await validateToken();
      if (!isValid) {
        throw new Error('Token has expired and refresh failed. Please run "node authenticate.js" to sign in again.');
      }
      console.error('Token still valid despite expiry time');
    }
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