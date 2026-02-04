import { Client } from '@microsoft/microsoft-graph-client';
import { PublicClientApplication } from '@azure/msal-node';
import { loadToken, saveToken, loadRefreshToken, saveRefreshToken, saveAccountInfo, isTokenExpired, isValidTokenFormat } from './token-store.js';
import { msalConfig } from './config.js';

// ============================================
// Constants
// ============================================

const READ_SCOPES: string[] = ['Notes.Read.All', 'User.Read', 'offline_access'];

// ============================================
// State Management
// ============================================

const pca = new PublicClientApplication(msalConfig);

let accessToken: string | null = null;
let tokenExpiresAt: Date | null = null;
let graphClient: Client | null = null;

// ============================================
// Authentication Functions
// ============================================

/**
 * Get the current access token (for direct API calls)
 */
export function getAccessToken(): string | null {
  return accessToken;
}

/**
 * Get the current Graph client instance
 */
export function getGraphClient(): Client | null {
  return graphClient;
}

/**
 * Validate token by making a simple API call
 */
async function validateToken(): Promise<boolean> {
  if (!accessToken) return false;
  if (!isValidTokenFormat(accessToken)) return false;
  
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    return response.ok;
  } catch {
    return false;
  }
}

/**
 * Refresh the access token using the stored refresh token
 */
async function refreshAccessToken(): Promise<boolean> {
  const refreshToken = await loadRefreshToken();
  if (!refreshToken) {
    console.error('No refresh token available');
    return false;
  }
  
  try {
    console.error('Attempting to refresh access token...');
    
    const response = await pca.acquireTokenByRefreshToken({
      refreshToken,
      scopes: READ_SCOPES
    });
    
    if (!response) {
      console.error('Token refresh returned null');
      return false;
    }
    
    accessToken = response.accessToken;
    tokenExpiresAt = response.expiresOn || new Date(Date.now() + 3600 * 1000);
    
    await saveToken(accessToken, tokenExpiresAt);
    
    if ((response as { refreshToken?: string }).refreshToken) {
      await saveRefreshToken((response as { refreshToken?: string }).refreshToken);
    }
    
    if (response.account) {
      await saveAccountInfo(response.account);
    }
    
    console.error('Token refreshed successfully, expires:', tokenExpiresAt);
    return true;
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    console.error('Failed to refresh token:', message);
    return false;
  }
}

/**
 * Ensure Graph client is created with valid token
 */
export async function ensureGraphClient(): Promise<Client> {
  if (!accessToken) {
    const stored = await loadToken();
    accessToken = stored.token;
    tokenExpiresAt = stored.expiresAt;
    console.error('Loaded token from store, expires:', tokenExpiresAt);
  }
  
  if (!accessToken || !isValidTokenFormat(accessToken)) {
    const refreshed = await refreshAccessToken();
    if (!refreshed) {
      throw new Error('No valid token found. Please run "npm run auth" first to sign in.');
    }
  }
  
  if (isTokenExpired(tokenExpiresAt)) {
    console.error('Token appears expired, attempting refresh...');
    const refreshed = await refreshAccessToken();
    
    if (!refreshed) {
      console.error('Refresh failed, validating current token with API...');
      const isValid = await validateToken();
      if (!isValid) {
        throw new Error('Token has expired and refresh failed. Please run "npm run auth" to sign in again.');
      }
      console.error('Token still valid despite expiry time');
    }
  }
  
  graphClient = Client.init({
    authProvider: (done) => done(null, accessToken)
  });
  
  return graphClient;
}
