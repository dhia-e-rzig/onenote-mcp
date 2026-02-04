import { describe, test, expect, vi, beforeEach } from 'vitest';
import { handleListSections, handleGetSection, handleCreateSection, handleListSectionGroups } from '../handlers/sections.js';

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

import { getGraphClient } from '../lib/graph-client.js';

describe('Section Handlers', () => {
  const mockApi = vi.fn();
  const mockGet = vi.fn();
  const mockPost = vi.fn();
  
  beforeEach(() => {
    vi.clearAllMocks();
    
    mockApi.mockReturnValue({
      get: mockGet,
      post: mockPost
    });
    
    (getGraphClient as ReturnType<typeof vi.fn>).mockReturnValue({
      api: mockApi
    });
  });

  describe('handleListSections', () => {
    test('returns all sections when no notebookId provided', async () => {
      const mockSections = [
        { id: 's1', displayName: 'Section 1' },
        { id: 's2', displayName: 'Section 2' }
      ];
      mockGet.mockResolvedValue({ value: mockSections });
      
      const result = await handleListSections();
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/sections');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockSections);
    });

    test('returns sections for specific notebook when notebookId provided', async () => {
      const mockSections = [{ id: 's1', displayName: 'Section 1' }];
      mockGet.mockResolvedValue({ value: mockSections });
      
      const result = await handleListSections('nb1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks/nb1/sections');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockSections);
    });

    test('returns error response on API failure', async () => {
      mockGet.mockRejectedValue(new Error('API Error'));
      
      const result = await handleListSections();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
      expect(parsed.error).toContain('List sections');
    });
  });

  describe('handleGetSection', () => {
    test('returns section when valid ID provided', async () => {
      const mockSection = { id: 's1', displayName: 'My Section' };
      mockGet.mockResolvedValue(mockSection);
      
      const result = await handleGetSection('s1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/sections/s1');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockSection);
    });

    test('returns error when sectionId is empty', async () => {
      const result = await handleGetSection('');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('sectionId is required');
      expect(mockGet).not.toHaveBeenCalled();
    });

    test('returns error response on API failure', async () => {
      mockGet.mockRejectedValue(new Error('Not found'));
      
      const result = await handleGetSection('invalid');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });

  describe('handleCreateSection', () => {
    test('creates section with valid parameters', async () => {
      const mockResponse = {
        id: 'new-s',
        displayName: 'New Section',
        self: 'https://...'
      };
      mockPost.mockResolvedValue(mockResponse);
      
      const result = await handleCreateSection('nb1', 'New Section');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks/nb1/sections');
      expect(mockPost).toHaveBeenCalledWith({ displayName: 'New Section' });
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.success).toBe(true);
      expect(parsed.id).toBe('new-s');
    });

    test('returns error when notebookId is empty', async () => {
      const result = await handleCreateSection('', 'Section Name');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toContain('notebookId is required');
      expect(mockPost).not.toHaveBeenCalled();
    });

    test('returns error when displayName is empty', async () => {
      const result = await handleCreateSection('nb1', '');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBe('displayName is required');
      expect(mockPost).not.toHaveBeenCalled();
    });

    test('returns error response on API failure', async () => {
      mockPost.mockRejectedValue(new Error('Create failed'));
      
      const result = await handleCreateSection('nb1', 'Test');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });

  describe('handleListSectionGroups', () => {
    test('returns all section groups when no notebookId provided', async () => {
      const mockGroups = [{ id: 'sg1', displayName: 'Group 1' }];
      mockGet.mockResolvedValue({ value: mockGroups });
      
      const result = await handleListSectionGroups();
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/sectionGroups');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockGroups);
    });

    test('returns section groups for specific notebook', async () => {
      const mockGroups = [{ id: 'sg1', displayName: 'Group 1' }];
      mockGet.mockResolvedValue({ value: mockGroups });
      
      const result = await handleListSectionGroups('nb1');
      
      expect(mockApi).toHaveBeenCalledWith('/me/onenote/notebooks/nb1/sectionGroups');
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed).toEqual(mockGroups);
    });

    test('returns error response on API failure', async () => {
      mockGet.mockRejectedValue(new Error('API Error'));
      
      const result = await handleListSectionGroups();
      
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.error).toBeDefined();
    });
  });
});
