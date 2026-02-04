import { createSafeErrorMessage } from '../lib/validation.js';

/**
 * Standard response type for all tool handlers
 */
export interface ToolResponse {
  [key: string]: unknown;
  content: Array<{ type: 'text'; text: string }>;
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
 * Create an error response with a safe error message
 */
export function errorResponse(operation: string, error: unknown): ToolResponse {
  const safeMessage = createSafeErrorMessage(operation, error);
  return {
    content: [{ type: 'text' as const, text: JSON.stringify({ error: safeMessage }) }]
  };
}
