import { describe, test, expect, vi, beforeEach } from 'vitest';
import { handleListNotebooks, handleGetNotebook, handleCreateNotebook } from '../handlers/notebooks.js';

// Mock dependencies
vi.mock('../lib/graph-client.js', () => ({
  ensureGraphClient: vi.fn(),
  getGraphClient: vi.fn()
}));

vi.mock('../lib/rate-limiter.js', () => ({
  rateLimiter: {
    execute: vi.fn((fn) => fn())
  }
}));

import { ensureGraphClient, getGraphClient } from '../lib/graph-client.js';
import { rateLimiter } from '../lib/rate-limiter.js';

describe('Notebook Handlers', () => {
  const mockApi = vi.fn();
  const mockGet = vi.fn();
  const mockPost = vi.fn();
  
  beforeEach(() => {
    vi.clearAllMocks();
    
    // Setup mock chain: graphClient.api().get() / .post()
    mockApi.mockReturnValue({
      get: mockGet,
      post: mockPost
    });
    
    (getGraphClient as ReturnType<typeof vi.fn>).mockReturnValue({
      api: mockApi
    });
  });

  describe('handleListNotebooks', () => {
    test('returns list of notebooks on success', async () => {
      const mockNotebooks = [
        { id: 'nb1', displayName: 'Notebook 1' },
        { id: 'nb2', displayName: 'Notebook 2' }
      ];
      mockGet.mockResolvedValue({ value: mockNotebooks });
      
      const result = await handleListNotebooks();
      
      expect(ensureGraphClient).toHaveBeenCalled();
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockNotebooks);
    });

    test('returns empty array when no notebooks exist', async () => {
      mockGet.mockResolvedValue({ value: [] });
      
      const result = await handleListNotebooks();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual([]);
    });

    test('returns error response on API failure', async () => {
      mockGet.mockRejectedValue(new Error('API Error'));
      
      const result = await handleListNotebooks();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
      expect(parsed.error).toContain('List notebooks');
    });

    test('uses rate limiter for API call', async () => {
      mockGet.mockResolvedValue({ value: [] });
      
      await handleListNotebooks();
      
      expect(rateLimiter.execute).toHaveBeenCalled();
    });
  });

  describe('handleGetNotebook', () => {
    test('returns specific notebook when ID provided', async () => {
      const mockNotebook = { id: 'nb1', displayName: 'My Notebook' };
      mockGet.mockResolvedValue(mockNotebook);
      
      const result = await handleGetNotebook('nb1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks/nb1');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockNotebook);
    });

    test('returns first notebook when no ID provided', async () => {
      const mockNotebooks = [
        { id: 'nb1', displayName: 'First Notebook' },
        { id: 'nb2', displayName: 'Second Notebook' }
      ];
      mockGet.mockResolvedValue({ value: mockNotebooks });
      
      const result = await handleGetNotebook();
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockNotebooks[0]);
    });

    test('returns error when no notebooks found and no ID provided', async () => {
      mockGet.mockResolvedValue({ value: [] });
      
      const result = await handleGetNotebook();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('No notebooks found');
    });

    test('returns error response on API failure', async () => {
      mockGet.mockRejectedValue(new Error('Not found'));
      
      const result = await handleGetNotebook('invalid-id');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });

  describe('handleCreateNotebook', () => {
    test('creates notebook with valid displayName', async () => {
      const mockResponse = {
        id: 'new-nb',
        displayName: 'New Notebook',
        self: 'https://...',
        links: { oneNoteWebUrl: { href: 'https://...' } }
      };
      mockPost.mockResolvedValue(mockResponse);
      
      const result = await handleCreateNotebook('New Notebook');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks');
      expect(mockPost).toHaveBeenCalledWith({ displayName: 'New Notebook' });
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.id).toBe('new-nb');
      expect(parsed.displayName).toBe('New Notebook');
    });

    test('returns error when displayName is empty', async () => {
      const result = await handleCreateNotebook('');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('displayName is required');
      expect(mockPost).not.toHaveBeenCalled();
    });

    test('returns error response on API failure', async () => {
      mockPost.mockRejectedValue(new Error('Create failed'));
      
      const result = await handleCreateNotebook('Test');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });
});
