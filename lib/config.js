// Central configuration file for OneNote MCP
// All shared settings in one place for easier updates

// Azure AD / Microsoft Identity Configuration
export const clientId = '813d941f-92ac-4ac0-94a2-1e89b720e15b';
export const authority = 'https://login.microsoftonline.com/consumers';
export const redirectUri = 'http://localhost:8400';

// OAuth Scopes
// Notes.ReadWrite - Read and write notebooks
// User.Read - Get user profile info
// offline_access - Get refresh tokens for persistent auth
export const scopes = ['Notes.ReadWrite', 'User.Read', 'offline_access'];

// MSAL Configuration
export const msalConfig = {
  auth: {
    clientId: clientId,
    authority: authority
  }
};

// Keytar service name for credential storage
export const keytarService = 'onenote-mcp';

// Rate limiting configuration
export const rateLimitConfig = {
  maxRequestsPerMinute: 60,
  retryDelayMs: 1000
};
