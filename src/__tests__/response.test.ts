import { describe, test, expect } from 'vitest';
import { successResponse, errorResponse } from '../handlers/response.js';

describe('Response Helpers', () => {
  describe('successResponse', () => {
    test('returns correct structure with string data', () => {
      const result = successResponse('test data');
      
      expect(result.content).toHaveLength(1);
      expect(result.content[0].type).toBe('text');
      expect(result.content[0].text).toBe('"test data"');
    });

    test('returns correct structure with object data', () => {
      const data = { id: '123', name: 'Test' };
      const result = successResponse(data);
      
      expect(result.content).toHaveLength(1);
      expect(result.content[0].type).toBe('text');
      expect(JSON.parse(result.content[0].text)).toEqual(data);
    });

    test('returns correct structure with array data', () => {
      const data = [{ id: '1' }, { id: '2' }];
      const result = successResponse(data);
      
      expect(result.content).toHaveLength(1);
      expect(JSON.parse(result.content[0].text)).toEqual(data);
    });

    test('returns correct structure with null', () => {
      const result = successResponse(null);
      
      expect(result.content[0].text).toBe('null');
    });

    test('has index signature compatibility', () => {
      const result = successResponse({ test: true });
      
      // Should be able to access content property
      expect(result.content).toBeDefined();
      // Index signature allows unknown keys
      expect(result['anyKey']).toBeUndefined();
    });
  });

  describe('errorResponse', () => {
    test('returns error structure with Error object', () => {
      const error = new Error('Something went wrong');
      const result = errorResponse('Test operation', error);
      
      expect(result.content).toHaveLength(1);
      expect(result.content[0].type).toBe('text');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toContain('Test operation');
    });

    test('returns error structure with string error', () => {
      const result = errorResponse('Test operation', 'string error');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });

    test('returns error structure with unknown error type', () => {
      const result = errorResponse('Test operation', { custom: 'error' });
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });

    test('sanitizes sensitive information from error', () => {
      const error = new Error('Token abc123xyz expired');
      const result = errorResponse('Auth', error);
      
      const parsed = JSON.parse(result.content[0].text);
      // Should not contain the raw token
      expect(parsed.error).not.toContain('abc123xyz');
    });
  });
});
