/**
 * Type definitions for OneNote MCP Server
 */

import type { AccountInfo } from '@azure/msal-node';

// Token-related types
export interface TokenData {
  token: string | null;
  expiresAt: Date | null;
}

export interface StoredTokenData {
  token: string;
  expiresAt: string | null;
}

export interface StoredAccountInfo {
  homeAccountId: string;
  environment: string;
  tenantId: string;
  username: string;
  localAccountId: string;
}

// Validation result types
export interface ValidationResult {
  valid: boolean;
  value: string;
  error?: string;
}

// Rate limiter types
export interface RateLimiterOptions {
  minDelay?: number;
  maxDelay?: number;
  maxRetries?: number;
}

// OneNote API response types
export interface OneNoteNotebook {
  id: string;
  displayName: string;
  self?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  links?: {
    oneNoteClientUrl?: { href: string };
    oneNoteWebUrl?: { href: string };
  };
}

export interface OneNoteSection {
  id: string;
  displayName: string;
  self?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  parentNotebook?: {
    id: string;
    displayName?: string;
  };
}

export interface OneNoteSectionGroup {
  id: string;
  displayName: string;
  self?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

export interface OneNotePage {
  id: string;
  title: string;
  self?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  contentUrl?: string;
  sectionName?: string;
  sectionId?: string;
  parentSection?: {
    id: string;
    displayName?: string;
  };
  links?: {
    oneNoteClientUrl?: { href: string };
    oneNoteWebUrl?: { href: string };
  };
}

export interface OneNotePageContent {
  id: string;
  title: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  contentUrl?: string;
  content: string;
}

export interface OneNoteListResponse<T> {
  value: T[];
  '@odata.context'?: string;
  '@odata.nextLink'?: string;
}

// MCP Tool response types
export interface McpToolResponse {
  content: Array<{
    type: 'text';
    text: string;
  }>;
}

// MSAL-related types
export interface PkceCodes {
  verifier: string;
  challenge: string;
}

// Graph API patch content
export interface PagePatchContent {
  target: string;
  action: 'append' | 'prepend' | 'replace';
  content: string;
}

// Config types
export interface MsalConfig {
  auth: {
    clientId: string;
    authority: string;
  };
}

export interface RateLimitConfig {
  maxRequestsPerMinute: number;
  retryDelayMs: number;
}

// Re-export AccountInfo for convenience
export type { AccountInfo };
