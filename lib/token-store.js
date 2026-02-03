/**
 * Secure Token Store Module
 * Uses OS credential manager via keytar for secure token storage
 */

import keytar from 'keytar';

// Credential store constants
const SERVICE_NAME = 'onenote-mcp';
const ACCOUNT_NAME = 'microsoft-graph-token';

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
 * Delete token from secure credential store
 */
export async function deleteToken() {
  await keytar.deletePassword(SERVICE_NAME, ACCOUNT_NAME);
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
 * @param {string} token - The access token
 * @returns {boolean}
 */
export function isValidTokenFormat(token) {
  if (!token || typeof token !== 'string') return false;
  // JWT tokens have 3 parts separated by dots
  const parts = token.split('.');
  return parts.length === 3 && parts.every(part => part.length > 0);
}
