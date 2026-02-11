import { createDetailedError, type DetailedError } from '../lib/validation.js';

/**
 * Standard response type for all tool handlers
 */
export interface ToolResponse {
  [key: string]: unknown;
  content: Array<{ type: 'text'; text: string }>;
}

/**
 * Error response details returned to the LLM
 */
interface ErrorResponseData {
  error: string;
  errorCode?: string;
  httpStatus?: number;
  timestamp: string;
  suggestion?: string;
  context?: Record<string, unknown>;
}

/**
 * Get a suggestion based on the error type
 */
function getSuggestion(details: DetailedError, context?: Record<string, unknown>): string | undefined {
  const code = details.graphErrorCode?.toLowerCase() || '';
  const msg = details.userMessage.toLowerCase();
  
  if (msg.includes('authentication') || msg.includes('sign in') || code.includes('auth')) {
    return 'Try running the authentication flow again or check if your token has expired.';
  }
  if (msg.includes('not found') || code === 'itemnotfound' || code === 'resourcenotfound') {
    const resourceType = context?.resourceType || 'resource';
    return `Verify the ${resourceType} ID is correct. Use the list operation to find valid IDs.`;
  }
  if (msg.includes('rate limit')) {
    return 'Wait a few seconds before retrying the request.';
  }
  if (msg.includes('permission') || msg.includes('access denied')) {
    return 'Check that your account has access to this OneNote resource.';
  }
  if (msg.includes('network') || msg.includes('connection')) {
    return 'Check your internet connection and try again.';
  }
  if (msg.includes('read-only') || msg.includes('locked')) {
    return 'This resource is read-only or locked. Try a different notebook or section.';
  }
  if (context?.sectionId && msg.includes('section')) {
    return 'Use listSections to find valid section IDs, or create a new section first.';
  }
  if (context?.notebookId && msg.includes('notebook')) {
    return 'Use listNotebooks to find valid notebook IDs.';
  }
  return undefined;
}

/**
 * Create a success response with JSON data
 */
export function successResponse(data: unknown): ToolResponse {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(data) }]
  };
}

/**
 * Create an error response with detailed error information
 * @param operation - The operation that failed
 * @param error - The error object
 * @param context - Optional context about the operation (resource IDs, etc.)
 */
export function errorResponse(
  operation: string, 
  error: unknown, 
  context?: Record<string, unknown>
): ToolResponse {
  const details = createDetailedError(operation, error, context);
  
  // Log detailed error for debugging
  console.error(details.logMessage);
  
  // Build response for LLM
  const responseData: ErrorResponseData = {
    error: details.userMessage,
    timestamp: details.timestamp
  };
  
  if (details.graphErrorCode) {
    responseData.errorCode = details.graphErrorCode;
  }
  
  if (details.httpStatus) {
    responseData.httpStatus = details.httpStatus;
  }
  
  const suggestion = getSuggestion(details, context);
  if (suggestion) {
    responseData.suggestion = suggestion;
  }
  
  // Include sanitized context (remove sensitive data)
  if (context) {
    const safeContext: Record<string, unknown> = {};
    const safeKeys = ['notebookId', 'sectionId', 'pageId', 'sectionGroupId', 'resourceType', 'query'];
    for (const key of safeKeys) {
      if (context[key] !== undefined) {
        safeContext[key] = context[key];
      }
    }
    if (Object.keys(safeContext).length > 0) {
      responseData.context = safeContext;
    }
  }
  
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(responseData) }]
  };
}
