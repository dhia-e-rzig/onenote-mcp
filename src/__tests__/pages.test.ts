import { describe, test, expect, vi, beforeEach } from 'vitest';
import {
  handleListPages,
  handleGetPage,
  handleCreatePage,
  handleUpdatePage,
  handleDeletePage,
  handleSearchPages
} from '../handlers/pages.js';

// Mock dependencies
vi.mock('../lib/graph-client.js', () => ({
  ensureGraphClient: vi.fn(),
  getGraphClient: vi.fn(),
  getAccessToken: vi.fn(() => 'mock-token')
}));

vi.mock('../lib/rate-limiter.js', () => ({
  rateLimiter: {
    execute: vi.fn((fn) => fn())
  }
}));

vi.mock('../lib/validation.js', () => ({
  sanitizeHtmlContent: vi.fn((content) => content),
  createSafeErrorMessage: vi.fn((operation, error) => `${operation} failed: ${error instanceof Error ? error.message : 'Unknown error'}`),
  createDetailedError: vi.fn((operation, error, context) => ({
    operation,
    userMessage: `${operation} failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
    logMessage: `${operation} FAILED`,
    timestamp: new Date().toISOString()
  }))
}));

// Mock global fetch
const mockFetch = vi.fn();
global.fetch = mockFetch;

import { getGraphClient } from '../lib/graph-client.js';
import { sanitizeHtmlContent } from '../lib/validation.js';

describe('Page Handlers', () => {
  const mockApi = vi.fn();
  const mockGet = vi.fn();
  const mockPost = vi.fn();
  const mockPatch = vi.fn();
  const mockDelete = vi.fn();
  const mockHeader = vi.fn();
  
  beforeEach(() => {
    vi.clearAllMocks();
    
    // Setup mock chain with fluent API support
    mockHeader.mockReturnValue({
      post: mockPost,
      patch: mockPatch
    });
    
    mockApi.mockReturnValue({
      get: mockGet,
      post: mockPost,
      patch: mockPatch,
      delete: mockDelete,
      header: mockHeader
    });
    
    (getGraphClient as ReturnType<typeof vi.fn>).mockReturnValue({
      api: mockApi
    });
  });

  describe('handleListPages', () => {
    test('returns pages for specific section when sectionId provided', async () => {
      const mockPages = [
        { id: 'p1', title: 'Page 1' },
        { id: 'p2', title: 'Page 2' }
      ];
      mockGet.mockResolvedValue({ value: mockPages });
      
      const result = await handleListPages('s1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/sections/s1/pages');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockPages);
    });

    test('returns pages from all sections when no sectionId provided', async () => {
      const mockSections = [
        { id: 's1', displayName: 'Section 1' },
        { id: 's2', displayName: 'Section 2' }
      ];
      const mockPagesS1 = [{ id: 'p1', title: 'Page 1' }];
      const mockPagesS2 = [{ id: 'p2', title: 'Page 2' }];
      
      mockGet
        .mockResolvedValueOnce({ value: mockSections })
        .mockResolvedValueOnce({ value: mockPagesS1 })
        .mockResolvedValueOnce({ value: mockPagesS2 });
      
      const result = await handleListPages();
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/sections');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toHaveLength(2);
      expect(parsed[0].sectionName).toBe('Section 1');
      expect(parsed[1].sectionName).toBe('Section 2');
    });

    test('returns empty array when no sections exist', async () => {
      mockGet.mockResolvedValue({ value: [] });
      
      const result = await handleListPages();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual([]);
    });

    test('returns error response on API failure', async () => {
      mockGet.mockRejectedValue(new Error('API Error'));
      
      const result = await handleListPages('s1');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });

  describe('handleGetPage', () => {
    test('returns page by ID with content', async () => {
      const mockPage = { id: 'p1', title: 'My Page', contentUrl: 'https://...' };
      mockGet.mockResolvedValue(mockPage);
      mockFetch.mockResolvedValue({
        ok: true,
        text: () => Promise.resolve('<html><body>Page content</body></html>')
      });
      
      const result = await handleGetPage('p1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/pages/p1');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe('p1');
      expect(parsed.title).toBe('My Page');
      expect(parsed.content).toContain('Page content');
    });

    test('returns page by title search', async () => {
      const mockPages = [
        { id: 'p1', title: 'My Page' },
        { id: 'p2', title: 'Other Page' }
      ];
      mockGet.mockResolvedValue({ value: mockPages });
      mockFetch.mockResolvedValue({
        ok: true,
        text: () => Promise.resolve('<html><body>Content</body></html>')
      });
      
      const result = await handleGetPage(undefined, 'My Page');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/pages');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.id).toBe('p1');
    });

    test('returns error when page not found by ID', async () => {
      mockGet.mockRejectedValue(new Error('Not found'));
      
      const result = await handleGetPage('invalid');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toContain('Page not found');
    });

    test('returns error when page not found by title', async () => {
      mockGet.mockResolvedValue({ value: [] });
      
      const result = await handleGetPage(undefined, 'Nonexistent');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toContain('No page found matching title');
    });

    test('returns error when neither pageId nor title provided', async () => {
      const result = await handleGetPage();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('Please provide either pageId or title parameter');
    });
  });

  describe('handleCreatePage', () => {
    test('creates page with all parameters', async () => {
      const mockResponse = {
        id: 'new-p',
        title: 'New Page',
        createdDateTime: '2024-01-01',
        self: 'https://...'
      };
      mockPost.mockResolvedValue(mockResponse);
      
      const result = await handleCreatePage('s1', 'New Page', '<p>Content</p>');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/sections/s1/pages');
      expect(mockHeader).toHaveBeenCalledWith('Content-Type', 'application/xhtml+xml');
      expect(sanitizeHtmlContent).toHaveBeenCalled();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.id).toBe('new-p');
    });

    test('creates page with default title and content', async () => {
      const mockResponse = { id: 'new-p', title: 'New Page' };
      mockPost.mockResolvedValue(mockResponse);
      
      const result = await handleCreatePage('s1');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
    });

    test('returns error when sectionId is empty', async () => {
      const result = await handleCreatePage('');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toContain('sectionId is required');
      expect(mockPost).not.toHaveBeenCalled();
    });

    test('returns error response on API failure', async () => {
      mockPost.mockRejectedValue(new Error('Create failed'));
      
      const result = await handleCreatePage('s1', 'Test');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });

  describe('handleUpdatePage', () => {
    test('updates page with content', async () => {
      mockPatch.mockResolvedValue({});
      
      const result = await handleUpdatePage('p1', '<p>New content</p>');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/pages/p1/content');
      expect(mockHeader).toHaveBeenCalledWith('Content-Type', 'application/json');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
    });

    test('updates page with custom target', async () => {
      mockPatch.mockResolvedValue({});
      
      const result = await handleUpdatePage('p1', '<p>Content</p>', 'div#custom');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
    });

    test('returns error when pageId is empty', async () => {
      const result = await handleUpdatePage('', '<p>Content</p>');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('pageId is required');
    });

    test('returns error when content is empty', async () => {
      const result = await handleUpdatePage('p1', '');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('content is required');
    });

    test('returns error response on API failure', async () => {
      mockPatch.mockRejectedValue(new Error('Update failed'));
      
      const result = await handleUpdatePage('p1', '<p>Content</p>');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });

  describe('handleDeletePage', () => {
    test('deletes page successfully', async () => {
      mockDelete.mockResolvedValue({});
      
      const result = await handleDeletePage('p1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/pages/p1');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toBe('Page deleted successfully');
    });

    test('returns error when pageId is empty', async () => {
      const result = await handleDeletePage('');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('pageId is required');
      expect(mockDelete).not.toHaveBeenCalled();
    });

    test('returns error response on API failure', async () => {
      mockDelete.mockRejectedValue(new Error('Delete failed'));
      
      const result = await handleDeletePage('p1');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });

  describe('handleSearchPages', () => {
    test('searches pages by title in all notebooks', async () => {
      const mockPages = [
        { id: 'p1', title: 'Test Page' },
        { id: 'p2', title: 'Other Page' }
      ];
      mockGet.mockResolvedValue({ value: mockPages });
      
      const result = await handleSearchPages('Test');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/pages');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.results).toHaveLength(1);
      expect(parsed.results[0].title).toBe('Test Page');
      expect(parsed.matches).toBe(1);
    });

    test('searches pages within specific section', async () => {
      const mockPages = [
        { id: 'p1', title: 'Test Page' },
        { id: 'p2', title: 'Another Test' }
      ];
      mockGet.mockResolvedValue({ value: mockPages });
      
      const result = await handleSearchPages('Test', undefined, 's1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/sections/s1/pages');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.results).toHaveLength(2);
      expect(parsed.matches).toBe(2);
    });

    test('searches pages within specific notebook', async () => {
      const mockSections = [{ id: 's1', displayName: 'Section 1' }];
      const mockPages = [{ id: 'p1', title: 'Test Page' }];
      
      mockGet
        .mockResolvedValueOnce({ value: mockSections })
        .mockResolvedValueOnce({ value: mockPages });
      
      const result = await handleSearchPages('Test', 'nb1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks/nb1/sections');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.results).toHaveLength(1);
      expect(parsed.matches).toBe(1);
    });

    test('returns empty array when no matches found', async () => {
      const mockPages = [{ id: 'p1', title: 'Other Page' }];
      mockGet.mockResolvedValue({ value: mockPages });
      
      const result = await handleSearchPages('Nonexistent');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.results).toEqual([]);
      expect(parsed.matches).toBe(0);
    });

    test('returns error when query is empty', async () => {
      const result = await handleSearchPages('');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('query parameter is required');
    });

    test('search is case-insensitive', async () => {
      const mockPages = [
        { id: 'p1', title: 'Test Page' },
        { id: 'p2', title: 'TEST PAGE' },
        { id: 'p3', title: 'test page' }
      ];
      mockGet.mockResolvedValue({ value: mockPages });
      
      const result = await handleSearchPages('test');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.results).toHaveLength(3);
      expect(parsed.matches).toBe(3);
    });
  });
});
