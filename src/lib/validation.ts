/**
 * Input Validation and Sanitization Module
 */

import type { ValidationResult } from '../types.js';

// Maximum lengths for inputs
const MAX_ID_LENGTH = 500;
const MAX_SEARCH_LENGTH = 200;
const MAX_TITLE_LENGTH = 200;

/**
 * Sanitize a string by removing potentially dangerous characters
 */
export function sanitizeString(input: string | null | undefined): string {
  if (!input || typeof input !== 'string') return '';
  // Remove null bytes and control characters except newlines and tabs
  // eslint-disable-next-line no-control-regex
  return input.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
}

/**
 * Validate and sanitize a page/notebook/section ID
 */
export function validateId(id: string | null | undefined): ValidationResult {
  if (!id || typeof id !== 'string') {
    return { valid: false, value: '', error: 'ID is required' };
  }
  
  const sanitized = sanitizeString(id).trim();
  
  if (sanitized.length === 0) {
    return { valid: false, value: '', error: 'ID cannot be empty' };
  }
  
  if (sanitized.length > MAX_ID_LENGTH) {
    return { valid: false, value: '', error: `ID exceeds maximum length of ${MAX_ID_LENGTH}` };
  }
  
  // OneNote IDs are typically alphanumeric with hyphens, underscores, and some special chars
  // Allow URL-safe characters
  if (!/^[a-zA-Z0-9\-_!.~'()*%]+$/.test(sanitized)) {
    return { valid: false, value: '', error: 'ID contains invalid characters' };
  }
  
  return { valid: true, value: sanitized };
}

/**
 * Validate and sanitize a search term
 */
export function validateSearchTerm(searchTerm: string | null | undefined): ValidationResult {
  if (!searchTerm || typeof searchTerm !== 'string') {
    return { valid: true, value: '' }; // Empty search is valid (returns all)
  }
  
  const sanitized = sanitizeString(searchTerm).trim();
  
  if (sanitized.length > MAX_SEARCH_LENGTH) {
    return { valid: false, value: '', error: `Search term exceeds maximum length of ${MAX_SEARCH_LENGTH}` };
  }
  
  return { valid: true, value: sanitized };
}

/**
 * Validate and sanitize a page title
 */
export function validateTitle(title: string | null | undefined): ValidationResult {
  if (!title || typeof title !== 'string') {
    return { valid: false, value: '', error: 'Title is required' };
  }
  
  const sanitized = sanitizeString(title).trim();
  
  if (sanitized.length === 0) {
    return { valid: false, value: '', error: 'Title cannot be empty' };
  }
  
  if (sanitized.length > MAX_TITLE_LENGTH) {
    return { valid: false, value: '', error: `Title exceeds maximum length of ${MAX_TITLE_LENGTH}` };
  }
  
  return { valid: true, value: sanitized };
}

/**
 * Sanitize HTML content for page creation
 */
export function sanitizeHtmlContent(html: string | null | undefined): string {
  if (!html || typeof html !== 'string') return '';
  
  // Remove script tags using a safer regex pattern (avoids ReDoS)
  // First, remove script tags with content
  let sanitized = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
  
  // Remove self-closing script tags
  sanitized = sanitized.replace(/<script[^>]*\/>/gi, '');
  
  // Remove event handlers (onclick, onload, etc.)
  sanitized = sanitized.replace(/\s+on\w+\s*=\s*"[^"]*"/gi, '');
  sanitized = sanitized.replace(/\s+on\w+\s*=\s*'[^']*'/gi, '');
  sanitized = sanitized.replace(/\s+on\w+\s*=\s*[^\s>]+/gi, '');
  
  return sanitized;
}

/**
 * Create a safe error message (no internal details)
 */
export function createSafeErrorMessage(operation: string, error: Error | unknown): string {
  // List of safe error messages to pass through
  const safeMessages = [
    'Page not found',
    'Notebook not found',
    'Section not found',
    'No sections found',
    'No notebooks found',
    'No pages found',
    'Authentication required',
    'Access denied',
    'Rate limit exceeded'
  ];
  
  const errorMessage = error instanceof Error ? error.message : '';
  
  // Check if it's a safe message
  for (const safe of safeMessages) {
    if (errorMessage.toLowerCase().includes(safe.toLowerCase())) {
      return `${operation}: ${safe}`;
    }
  }
  
  // Check for HTTP status codes
  const httpMatch = errorMessage.match(/status[:\s]*(\d{3})/i);
  if (httpMatch) {
    const status = parseInt(httpMatch[1], 10);
    if (status === 401) return `${operation}: Authentication required`;
    if (status === 403) return `${operation}: Access denied`;
    if (status === 404) return `${operation}: Resource not found`;
    if (status === 429) return `${operation}: Rate limit exceeded, please try again later`;
    if (status >= 500) return `${operation}: Service temporarily unavailable`;
  }
  
  // Generic error for unknown cases
  return `${operation}: Operation failed. Please try again.`;
}
