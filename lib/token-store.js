/**
 * Secure Token Store Module
 * Uses OS credential manager via keytar for secure token storage
 * Supports refresh tokens for persistent authentication across sessions
 */

import keytar from 'keytar';
import { keytarService } from './config.js';

// Credential store constants
const SERVICE_NAME = keytarService;
const ACCOUNT_NAME = 'microsoft-graph-token';
const REFRESH_ACCOUNT_NAME = 'microsoft-graph-refresh-token';
const ACCOUNT_INFO_NAME = 'microsoft-graph-account-info';

/**
 * Load token from secure credential store
 * @returns {Promise<{token: string|null, expiresAt: Date|null}>}
 */
export async function loadToken() {
  try {
    const tokenData = await keytar.getPassword(SERVICE_NAME, ACCOUNT_NAME);
    if (tokenData) {
      try {
        const parsed = JSON.parse(tokenData);
        return {
          token: parsed.token,
          expiresAt: parsed.expiresAt ? new Date(parsed.expiresAt) : null
        };
      } catch (parseError) {
        // Legacy format - raw token
        return { token: tokenData, expiresAt: null };
      }
    }
  } catch (error) {
    console.error('Error reading token from credential store:', error.message);
  }
  return { token: null, expiresAt: null };
}

/**
 * Load refresh token from secure credential store
 * @returns {Promise<string|null>}
 */
export async function loadRefreshToken() {
  try {
    return await keytar.getPassword(SERVICE_NAME, REFRESH_ACCOUNT_NAME);
  } catch (error) {
    console.error('Error reading refresh token from credential store:', error.message);
  }
  return null;
}

/**
 * Load account info for silent token acquisition
 * @returns {Promise<object|null>}
 */
export async function loadAccountInfo() {
  try {
    const accountData = await keytar.getPassword(SERVICE_NAME, ACCOUNT_INFO_NAME);
    if (accountData) {
      return JSON.parse(accountData);
    }
  } catch (error) {
    console.error('Error reading account info from credential store:', error.message);
  }
  return null;
}

/**
 * Save token to secure credential store
 * @param {string} token - The access token
 * @param {Date|null} expiresAt - Token expiration date
 */
export async function saveToken(token, expiresAt) {
  const tokenData = JSON.stringify({ 
    token: token, 
    expiresAt: expiresAt ? expiresAt.toISOString() : null 
  });
  await keytar.setPassword(SERVICE_NAME, ACCOUNT_NAME, tokenData);
}

/**
 * Save refresh token to secure credential store
 * @param {string} refreshToken - The refresh token
 */
export async function saveRefreshToken(refreshToken) {
  if (refreshToken) {
    await keytar.setPassword(SERVICE_NAME, REFRESH_ACCOUNT_NAME, refreshToken);
  }
}

/**
 * Save account info for silent token acquisition
 * @param {object} account - The MSAL account object
 */
export async function saveAccountInfo(account) {
  if (account) {
    const accountData = JSON.stringify({
      homeAccountId: account.homeAccountId,
      environment: account.environment,
      tenantId: account.tenantId,
      username: account.username,
      localAccountId: account.localAccountId
    });
    await keytar.setPassword(SERVICE_NAME, ACCOUNT_INFO_NAME, accountData);
  }
}

/**
 * Delete all tokens from secure credential store
 */
export async function deleteToken() {
  await keytar.deletePassword(SERVICE_NAME, ACCOUNT_NAME);
  await keytar.deletePassword(SERVICE_NAME, REFRESH_ACCOUNT_NAME);
  await keytar.deletePassword(SERVICE_NAME, ACCOUNT_INFO_NAME);
}

/**
 * Check if token is expired or about to expire (within 5 minutes)
 * @param {Date|null} expiresAt - Token expiration date
 * @returns {boolean}
 */
export function isTokenExpired(expiresAt) {
  if (!expiresAt) return true;
  const bufferTime = 5 * 60 * 1000; // 5 minutes
  return new Date().getTime() > (expiresAt.getTime() - bufferTime);
}

/**
 * Validate token format (basic check)
 * Microsoft Graph can return both JWT tokens (3 parts) and opaque tokens
 * @param {string} token - The access token
 * @returns {boolean}
 */
export function isValidTokenFormat(token) {
  if (!token || typeof token !== 'string') return false;
  // Microsoft tokens can be:
  // 1. JWT tokens (3 parts separated by dots)
  // 2. Opaque tokens (single string, typically starting with "EwB" or similar)
  // We accept both formats as long as they have reasonable length
  if (token.length < 10) return false;
  
  const parts = token.split('.');
  // JWT format (3 parts)
  if (parts.length === 3 && parts.every(part => part.length > 0)) {
    return true;
  }
  // Opaque token format (no dots, reasonable length for Microsoft tokens)
  if (parts.length === 1 && token.length > 100) {
    return true;
  }
  return false;
}
