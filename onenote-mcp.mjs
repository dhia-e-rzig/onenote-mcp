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
import { clientId, authority, msalConfig } from './lib/config.js';

// Load environment variables
dotenv.config();

// Warn if using default client ID
if (!process.env.AZURE_CLIENT_ID) {
  console.error('WARNING: Using default Microsoft Graph Explorer client ID.');
  console.error('For production, register your own Azure AD application and set AZURE_CLIENT_ID environment variable.');
}

// OAuth scopes - use read-only by default, write only when needed
// Include offline_access for refresh tokens
const READ_SCOPES = ['Notes.Read.All', 'User.Read', 'offline_access'];
const WRITE_SCOPES = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read', 'offline_access'];

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
  "Get details of a specific notebook by ID",
  {
    notebookId: {
      type: "string",
      description: "The ID of the notebook to get. If not provided, returns the first notebook."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      if (params.notebookId) {
        const response = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/notebooks/${params.notebookId}`).get()
        );
        return { 
          content: [{ type: "text", text: JSON.stringify(response) }]
        };
      }
      
      // Fallback: return first notebook
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
  "List all sections. Provide a notebookId to list sections from a specific notebook, or omit to list all sections.",
  {
    notebookId: {
      type: "string",
      description: "The ID of the notebook to list sections from. If not provided, lists all sections across all notebooks."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      let response;
      if (params.notebookId) {
        // List sections from specific notebook
        response = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/notebooks/${params.notebookId}/sections`).get()
        );
      } else {
        // List all sections
        response = await rateLimiter.execute(() => 
          graphClient.api("/me/onenote/sections").get()
        );
      }
      
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

// Tool for getting section details
server.tool(
  "getSection",
  "Get details of a specific section by ID",
  {
    sectionId: {
      type: "string",
      description: "The ID of the section to get. Required."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      if (!params.sectionId) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'sectionId is required' }) }]
        };
      }
      
      const response = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/sections/${params.sectionId}`).get()
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify(response) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Get section', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for listing section groups
server.tool(
  "listSectionGroups",
  "List all section groups. Provide a notebookId to list section groups from a specific notebook.",
  {
    notebookId: {
      type: "string",
      description: "The ID of the notebook to list section groups from. If not provided, lists all section groups."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      let response;
      if (params.notebookId) {
        response = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/notebooks/${params.notebookId}/sectionGroups`).get()
        );
      } else {
        response = await rateLimiter.execute(() => 
          graphClient.api("/me/onenote/sectionGroups").get()
        );
      }
      
      return { 
        content: [{ type: "text", text: JSON.stringify(response.value) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('List section groups', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for listing pages in a section
server.tool(
  "listPages",
  "List all pages in a section. Provide a sectionId to list pages from a specific section, or omit to list pages from all sections.",
  {
    sectionId: {
      type: "string",
      description: "The ID of the section to list pages from. If not provided, lists pages from all sections."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      if (params.sectionId) {
        // List pages from specific section
        const response = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/sections/${params.sectionId}/pages`).get()
        );
        return { 
          content: [{ type: "text", text: JSON.stringify(response.value) }]
        };
      } else {
        // List pages from all sections (get sections first, then pages from each)
        const sectionsResponse = await rateLimiter.execute(() => 
          graphClient.api("/me/onenote/sections").get()
        );
        
        if (!sectionsResponse.value || sectionsResponse.value.length === 0) {
          return { 
            content: [{ type: "text", text: "[]" }]
          };
        }
        
        // Collect pages from all sections
        const allPages = [];
        for (const section of sectionsResponse.value) {
          try {
            const pagesResponse = await rateLimiter.execute(() => 
              graphClient.api(`/me/onenote/sections/${section.id}/pages`).get()
            );
            if (pagesResponse.value) {
              // Add section info to each page for context
              for (const page of pagesResponse.value) {
                page.sectionName = section.displayName;
                page.sectionId = section.id;
                allPages.push(page);
              }
            }
          } catch (e) {
            // Skip sections that fail (might have permission issues)
          }
        }
        
        return { 
          content: [{ type: "text", text: JSON.stringify(allPages) }]
        };
      }
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
  "Get the content of a page by ID or title",
  {
    pageId: {
      type: "string",
      description: "The ID of the page to retrieve. Takes precedence over title search."
    },
    title: {
      type: "string",
      description: "Search for a page by title (partial match). Used if pageId is not provided."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      let targetPage = null;
      
      if (params.pageId) {
        // Direct page ID lookup
        try {
          const pageInfo = await rateLimiter.execute(() => 
            graphClient.api(`/me/onenote/pages/${params.pageId}`).get()
          );
          targetPage = pageInfo;
        } catch (e) {
          return {
            content: [{ type: "text", text: JSON.stringify({ error: `Page not found with ID: ${params.pageId}` }) }]
          };
        }
      } else if (params.title) {
        // Search by title
        const pagesResponse = await rateLimiter.execute(() => 
          graphClient.api('/me/onenote/pages').get()
        );
        
        if (pagesResponse.value && pagesResponse.value.length > 0) {
          const searchLower = params.title.toLowerCase();
          targetPage = pagesResponse.value.find(p => 
            p.title && p.title.toLowerCase().includes(searchLower)
          );
        }
        
        if (!targetPage) {
          return {
            content: [{ type: "text", text: JSON.stringify({ error: `No page found matching title: ${params.title}` }) }]
          };
        }
      } else {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'Please provide either pageId or title parameter' }) }]
        };
      }
      
      // Get page content
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
        content: [{ 
          type: "text", 
          text: JSON.stringify({
            id: targetPage.id,
            title: targetPage.title,
            createdDateTime: targetPage.createdDateTime,
            lastModifiedDateTime: targetPage.lastModifiedDateTime,
            contentUrl: targetPage.contentUrl,
            content: content
          })
        }]
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
  "Create a new page in a specific section",
  {
    sectionId: {
      type: "string",
      description: "The ID of the section to create the page in. Required."
    },
    title: {
      type: "string",
      description: "The title of the new page. Defaults to 'New Page' if not provided."
    },
    content: {
      type: "string",
      description: "The HTML content for the page body. Can include basic HTML tags like <p>, <h1>, <ul>, etc."
    }
  },
  async (params) => {
    try {
      // Require write permissions for this operation
      await ensureGraphClient(true);
      
      if (!params.sectionId) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'sectionId is required. Use listSections to find section IDs.' }) }]
        };
      }
      
      const pageTitle = params.title || 'New Page';
      const bodyContent = params.content || '<p>This is a new page created via the OneNote MCP.</p>';
      
      // Sanitize HTML content
      const simpleHtml = sanitizeHtmlContent(`
        <!DOCTYPE html>
        <html>
          <head>
            <title>${pageTitle}</title>
          </head>
          <body>
            ${bodyContent}
          </body>
        </html>
      `);
      
      const response = await rateLimiter.execute(() => 
        graphClient
          .api(`/me/onenote/sections/${params.sectionId}/pages`)
          .header("Content-Type", "application/xhtml+xml")
          .post(simpleHtml)
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify({ 
          success: true, 
          id: response.id,
          title: response.title,
          createdDateTime: response.createdDateTime,
          self: response.self
        }) }]
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
  "Search for pages by title across all notebooks",
  {
    query: {
      type: "string",
      description: "The search term to find in page titles. Required."
    },
    notebookId: {
      type: "string",
      description: "Optional: Limit search to pages within a specific notebook."
    },
    sectionId: {
      type: "string",
      description: "Optional: Limit search to pages within a specific section."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(false);
      
      if (!params.query) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'query parameter is required' }) }]
        };
      }
      
      let pages = [];
      
      if (params.sectionId) {
        // Search within specific section
        const response = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/sections/${params.sectionId}/pages`).get()
        );
        pages = response.value || [];
      } else if (params.notebookId) {
        // Search within specific notebook (get all sections first)
        const sectionsResponse = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/notebooks/${params.notebookId}/sections`).get()
        );
        
        for (const section of sectionsResponse.value || []) {
          try {
            const pagesResponse = await rateLimiter.execute(() => 
              graphClient.api(`/me/onenote/sections/${section.id}/pages`).get()
            );
            if (pagesResponse.value) {
              for (const page of pagesResponse.value) {
                page.sectionName = section.displayName;
                page.sectionId = section.id;
                pages.push(page);
              }
            }
          } catch (e) { /* skip */ }
        }
      } else {
        // Search all pages
        const response = await rateLimiter.execute(() => 
          graphClient.api("/me/onenote/pages").get()
        );
        pages = response.value || [];
      }
      
      // Filter by search term
      const searchTerm = params.query.toLowerCase();
      const filteredPages = pages.filter(page => 
        page.title && page.title.toLowerCase().includes(searchTerm)
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify(filteredPages) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Search pages', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for deleting a page
server.tool(
  "deletePage",
  "Delete a page by ID",
  {
    pageId: {
      type: "string",
      description: "The ID of the page to delete. Required."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(true); // Require write permissions
      
      if (!params.pageId) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'pageId is required' }) }]
        };
      }
      
      await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/pages/${params.pageId}`).delete()
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify({ success: true, message: 'Page deleted successfully' }) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Delete page', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for updating/appending content to a page
server.tool(
  "updatePage",
  "Update a page by appending content to it",
  {
    pageId: {
      type: "string",
      description: "The ID of the page to update. Required."
    },
    content: {
      type: "string", 
      description: "The HTML content to append to the page. Required."
    },
    target: {
      type: "string",
      description: "Where to insert the content: 'body' (default) appends to page body, or specify an element ID."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(true); // Require write permissions
      
      if (!params.pageId) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'pageId is required' }) }]
        };
      }
      
      if (!params.content) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'content is required' }) }]
        };
      }
      
      const target = params.target || 'body';
      
      // OneNote PATCH requires a specific format
      const patchContent = [
        {
          target: target,
          action: 'append',
          content: sanitizeHtmlContent(params.content)
        }
      ];
      
      await rateLimiter.execute(() => 
        graphClient
          .api(`/me/onenote/pages/${params.pageId}/content`)
          .header('Content-Type', 'application/json')
          .patch(patchContent)
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify({ success: true, message: 'Page updated successfully' }) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Update page', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for creating a new section in a notebook
server.tool(
  "createSection",
  "Create a new section in a notebook",
  {
    notebookId: {
      type: "string",
      description: "The ID of the notebook to create the section in. Required."
    },
    displayName: {
      type: "string",
      description: "The name of the new section. Required."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(true); // Require write permissions
      
      if (!params.notebookId) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'notebookId is required. Use listNotebooks to find notebook IDs.' }) }]
        };
      }
      
      if (!params.displayName) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'displayName is required' }) }]
        };
      }
      
      const response = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/notebooks/${params.notebookId}/sections`).post({
          displayName: params.displayName
        })
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify({ 
          success: true, 
          id: response.id,
          displayName: response.displayName,
          self: response.self
        }) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Create section', error);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: safeMessage }) }]
      };
    }
  }
);

// Tool for creating a new notebook
server.tool(
  "createNotebook",
  "Create a new notebook",
  {
    displayName: {
      type: "string",
      description: "The name of the new notebook. Required."
    }
  },
  async (params) => {
    try {
      await ensureGraphClient(true); // Require write permissions
      
      if (!params.displayName) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: 'displayName is required' }) }]
        };
      }
      
      const response = await rateLimiter.execute(() => 
        graphClient.api('/me/onenote/notebooks').post({
          displayName: params.displayName
        })
      );
      
      return { 
        content: [{ type: "text", text: JSON.stringify({ 
          success: true, 
          id: response.id,
          displayName: response.displayName,
          self: response.self,
          links: response.links
        }) }]
      };
    } catch (error) {
      const safeMessage = createSafeErrorMessage('Create notebook', error);
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