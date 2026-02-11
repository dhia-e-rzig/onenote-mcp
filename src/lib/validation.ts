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
 * Detailed error information for logging and response
 */
export interface DetailedError {
  operation: string;
  userMessage: string;
  logMessage: string;
  httpStatus?: number;
  graphErrorCode?: string;
  requestId?: string;
  timestamp: string;
}

/**
 * Extract Graph API error details from error object
 */
function extractGraphApiDetails(error: unknown): { 
  code?: string; 
  message?: string; 
  status?: number;
  requestId?: string;
  innerError?: unknown;
} {
  if (!error || typeof error !== 'object') return {};
  
  const err = error as Record<string, unknown>;
  
  // Microsoft Graph SDK errors often have body.error structure
  if (err.body && typeof err.body === 'object') {
    const body = err.body as Record<string, unknown>;
    if (body.error && typeof body.error === 'object') {
      const graphError = body.error as Record<string, unknown>;
      return {
        code: graphError.code as string | undefined,
        message: graphError.message as string | undefined,
        innerError: graphError.innerError,
        requestId: (graphError.innerError as Record<string, unknown> | undefined)?.['request-id'] as string | undefined
      };
    }
  }
  
  // Direct error structure
  if (err.code || err.message) {
    return {
      code: err.code as string | undefined,
      message: err.message as string | undefined,
      status: err.statusCode as number | undefined || err.status as number | undefined
    };
  }
  
  return {};
}

/**
 * Map Graph API error codes to user-friendly messages
 */
function mapGraphErrorCode(code: string | undefined): string | null {
  if (!code) return null;
  
  const errorCodeMap: Record<string, string> = {
    'InvalidAuthenticationToken': 'Authentication token is invalid or expired. Please re-authenticate.',
    'AuthenticationError': 'Authentication failed. Please check your credentials and try again.',
    'Authorization_RequestDenied': 'Access denied. You may not have permission for this resource.',
    'AccessDenied': 'Access denied to this resource. Check your permissions.',
    'ItemNotFound': 'The requested item was not found. It may have been deleted or moved.',
    'ResourceNotFound': 'The requested resource does not exist.',
    'NotFound': 'Resource not found.',
    'BadRequest': 'Invalid request. Please check your parameters.',
    'InvalidRequest': 'The request was invalid. Verify the input format.',
    'MalformedRequest': 'The request format is incorrect.',
    'QuotaExceeded': 'API quota exceeded. Please wait before making more requests.',
    'TooManyRequests': 'Rate limit exceeded. Please wait and try again.',
    'ServiceNotAvailable': 'OneNote service is temporarily unavailable. Try again later.',
    'GeneralException': 'An unexpected error occurred in the OneNote service.',
    'NameAlreadyExists': 'An item with this name already exists.',
    'NotSupported': 'This operation is not supported.',
    'Conflict': 'A conflict occurred. The resource may have been modified.',
    'PreconditionFailed': 'The resource has been modified since it was retrieved.',
    'InternalServerError': 'An internal server error occurred. Please try again.',
    'UnknownError': 'An unknown error occurred.',
    '20102': 'Invalid page content. Check your HTML format.',
    '20103': 'The page content is too large.',
    '20104': 'Unsupported media type in page content.',
    '20108': 'Invalid section or notebook reference.',
    '20115': 'The notebook is read-only.',
    '20117': 'The section is locked.',
    '20119': 'Cannot create content in a read-only section.',
    '20156': 'Invalid OneNote structure.'
  };
  
  return errorCodeMap[code] || null;
}

/**
 * Sanitize error message to remove sensitive data like tokens
 */
function sanitizeErrorMessage(message: string): string {
  // Remove potential tokens (JWT format: xxx.xxx.xxx)
  let sanitized = message.replace(/\b[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}\b/g, '[TOKEN_REDACTED]');
  
  // Remove potential bearer tokens
  sanitized = sanitized.replace(/Bearer\s+[A-Za-z0-9_\-.]+/gi, 'Bearer [TOKEN_REDACTED]');
  
  // Remove potential API keys or secrets (alphanumeric strings 20+ chars)
  sanitized = sanitized.replace(/(?:key|token|secret|password|credential)[=:\s]+['"]?[A-Za-z0-9_\-]{20,}['"]?/gi, '[CREDENTIAL_REDACTED]');
  
  // Remove shorter token-like values that follow "token" (captures "Token abc123xyz")
  sanitized = sanitized.replace(/\b(token)\s+([A-Za-z0-9_\-]{5,})\b/gi, '$1 [REDACTED]');
  
  return sanitized;
}

/**
 * Create detailed error information for both logging and response
 */
export function createDetailedError(operation: string, error: Error | unknown, context?: Record<string, unknown>): DetailedError {
  const timestamp = new Date().toISOString();
  const graphDetails = extractGraphApiDetails(error);
  const errorMessage = error instanceof Error ? error.message : String(error || 'Unknown error');
  
  // Extract HTTP status
  let httpStatus: number | undefined;
  const httpMatch = errorMessage.match(/status[:\s]*(\d{3})/i);
  if (httpMatch) {
    httpStatus = parseInt(httpMatch[1], 10);
  } else if (graphDetails.status) {
    httpStatus = graphDetails.status;
  }
  
  // Build detailed log message
  const logParts = [
    `[${timestamp}] ${operation} FAILED`,
    `Error: ${errorMessage}`,
  ];
  
  if (graphDetails.code) {
    logParts.push(`Graph Error Code: ${graphDetails.code}`);
  }
  if (graphDetails.message && graphDetails.message !== errorMessage) {
    logParts.push(`Graph Message: ${graphDetails.message}`);
  }
  if (graphDetails.requestId) {
    logParts.push(`Request ID: ${graphDetails.requestId}`);
  }
  if (httpStatus) {
    logParts.push(`HTTP Status: ${httpStatus}`);
  }
  if (context) {
    logParts.push(`Context: ${JSON.stringify(context)}`);
  }
  if (error instanceof Error && error.stack) {
    logParts.push(`Stack: ${error.stack.split('\n').slice(0, 3).join(' -> ')}`);
  }
  
  const logMessage = logParts.join(' | ');
  
  // Build user-friendly message
  let userMessage: string;
  
  // Try to map Graph error code first
  const mappedMessage = mapGraphErrorCode(graphDetails.code);
  if (mappedMessage) {
    userMessage = `${operation}: ${mappedMessage}`;
  } else if (graphDetails.message) {
    // Use Graph message if it's reasonably user-friendly, but sanitize it
    const cleanMessage = sanitizeErrorMessage(graphDetails.message.replace(/\s+/g, ' ').trim());
    if (cleanMessage.length < 200) {
      userMessage = `${operation}: ${cleanMessage}`;
    } else {
      userMessage = `${operation}: ${cleanMessage.substring(0, 197)}...`;
    }
  } else if (httpStatus) {
    // Map HTTP status codes
    const statusMessages: Record<number, string> = {
      400: 'Bad request. Please check your input parameters.',
      401: 'Authentication required. Please sign in again.',
      403: 'Access denied. You may not have permission for this operation.',
      404: 'Resource not found. It may have been deleted or the ID is incorrect.',
      409: 'Conflict. The resource may have been modified by another process.',
      429: 'Rate limit exceeded. Please wait a moment and try again.',
      500: 'OneNote service error. Please try again later.',
      502: 'Service temporarily unavailable. Please try again.',
      503: 'OneNote service is temporarily unavailable. Please try again later.',
      504: 'Request timed out. Please try again.'
    };
    userMessage = `${operation}: ${statusMessages[httpStatus] || `Server returned error ${httpStatus}`}`;
  } else {
    // Check for common error patterns in message
    const lowerMsg = errorMessage.toLowerCase();
    if (lowerMsg.includes('network') || lowerMsg.includes('enotfound') || lowerMsg.includes('econnrefused')) {
      userMessage = `${operation}: Network connection error. Please check your internet connection.`;
    } else if (lowerMsg.includes('timeout')) {
      userMessage = `${operation}: Request timed out. Please try again.`;
    } else if (lowerMsg.includes('authentication') || lowerMsg.includes('auth')) {
      userMessage = `${operation}: Authentication error. Please sign in again.`;
    } else {
      userMessage = `${operation}: Operation failed - ${sanitizeErrorMessage(errorMessage.substring(0, 100))}`;
    }
  }
  
  return {
    operation,
    userMessage,
    logMessage,
    httpStatus,
    graphErrorCode: graphDetails.code,
    requestId: graphDetails.requestId,
    timestamp
  };
}

/**
 * Create a safe error message (no internal details) - DEPRECATED: Use createDetailedError instead
 */
export function createSafeErrorMessage(operation: string, error: Error | unknown): string {
  return createDetailedError(operation, error).userMessage;
}
