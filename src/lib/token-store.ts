/**
 * Secure Token Store Module
 * Uses OS credential manager via keytar for secure token storage
 * Supports refresh tokens for persistent authentication across sessions
 */

import keytar from 'keytar';
import { keytarService } from './config.js';
import type { TokenData, StoredTokenData, StoredAccountInfo } from '../types.js';

// Credential store constants
const SERVICE_NAME: string = keytarService;
const ACCOUNT_NAME: string = 'microsoft-graph-token';
const REFRESH_ACCOUNT_NAME: string = 'microsoft-graph-refresh-token';
const ACCOUNT_INFO_NAME: string = 'microsoft-graph-account-info';

/**
 * Load token from secure credential store
 */
export async function loadToken(): Promise<TokenData> {
  try {
    const tokenData = await keytar.getPassword(SERVICE_NAME, ACCOUNT_NAME);
    if (tokenData) {
      try {
        const parsed: StoredTokenData = JSON.parse(tokenData);
        return {
          token: parsed.token,
          expiresAt: parsed.expiresAt ? new Date(parsed.expiresAt) : null
        };
      } catch {
        // Legacy format - raw token
        return { token: tokenData, expiresAt: null };
      }
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    console.error('Error reading token from credential store:', message);
  }
  return { token: null, expiresAt: null };
}

/**
 * Load refresh token from secure credential store
 */
export async function loadRefreshToken(): Promise<string | null> {
  try {
    return await keytar.getPassword(SERVICE_NAME, REFRESH_ACCOUNT_NAME);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    console.error('Error reading refresh token from credential store:', message);
  }
  return null;
}

/**
 * Load account info for silent token acquisition
 */
export async function loadAccountInfo(): Promise<StoredAccountInfo | null> {
  try {
    const accountData = await keytar.getPassword(SERVICE_NAME, ACCOUNT_INFO_NAME);
    if (accountData) {
      return JSON.parse(accountData) as StoredAccountInfo;
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    console.error('Error reading account info from credential store:', message);
  }
  return null;
}

/**
 * Save token to secure credential store
 */
export async function saveToken(token: string, expiresAt: Date | null): Promise<void> {
  const tokenData: StoredTokenData = { 
    token: token, 
    expiresAt: expiresAt ? expiresAt.toISOString() : null 
  };
  await keytar.setPassword(SERVICE_NAME, ACCOUNT_NAME, JSON.stringify(tokenData));
}

/**
 * Save refresh token to secure credential store
 */
export async function saveRefreshToken(refreshToken: string | undefined): Promise<void> {
  if (refreshToken) {
    await keytar.setPassword(SERVICE_NAME, REFRESH_ACCOUNT_NAME, refreshToken);
  }
}

/**
 * Save account info for silent token acquisition
 */
export async function saveAccountInfo(account: {
  homeAccountId?: string;
  environment?: string;
  tenantId?: string;
  username?: string;
  localAccountId?: string;
} | null): Promise<void> {
  if (account) {
    const accountData: StoredAccountInfo = {
      homeAccountId: account.homeAccountId || '',
      environment: account.environment || '',
      tenantId: account.tenantId || '',
      username: account.username || '',
      localAccountId: account.localAccountId || ''
    };
    await keytar.setPassword(SERVICE_NAME, ACCOUNT_INFO_NAME, JSON.stringify(accountData));
  }
}

/**
 * Delete all tokens from secure credential store
 */
export async function deleteToken(): Promise<void> {
  await keytar.deletePassword(SERVICE_NAME, ACCOUNT_NAME);
  await keytar.deletePassword(SERVICE_NAME, REFRESH_ACCOUNT_NAME);
  await keytar.deletePassword(SERVICE_NAME, ACCOUNT_INFO_NAME);
}

/**
 * Check if token is expired or about to expire (within 5 minutes)
 */
export function isTokenExpired(expiresAt: Date | null): boolean {
  if (!expiresAt) return true;
  const bufferTime = 5 * 60 * 1000; // 5 minutes
  return new Date().getTime() > (expiresAt.getTime() - bufferTime);
}

/**
 * Validate token format (basic check)
 * Microsoft Graph can return both JWT tokens (3 parts) and opaque tokens
 */
export function isValidTokenFormat(token: string | null | undefined): boolean {
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
