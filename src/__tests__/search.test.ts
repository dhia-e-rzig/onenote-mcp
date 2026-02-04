import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock the graph client
vi.mock('../lib/graph-client.js', () => ({
  ensureGraphClient: vi.fn().mockResolvedValue(undefined),
  getGraphClient: vi.fn()
}));

// Mock the rate limiter
vi.mock('../lib/rate-limiter.js', () => ({
  rateLimiter: {
    execute: vi.fn((fn) => fn())
  }
}));

import { getGraphClient } from '../lib/graph-client.js';
import {
  handleSearchNotebooks,
  handleSearchSections,
  handleSearchSectionGroups,
  handleUniversalSearch
} from '../handlers/search.js';

describe('Search Handlers', () => {
  const mockGraphClient = {
    api: vi.fn().mockReturnThis(),
    get: vi.fn()
  };

  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(getGraphClient).mockReturnValue(mockGraphClient as any);
  });

  describe('handleSearchNotebooks', () => {
    it('returns error when query is empty', async () => {
      const result = await handleSearchNotebooks('');
      const data = JSON.parse(result.content[0].text);
      expect(data.error).toBe('query parameter is required');
    });

    it('returns matching notebooks ranked by relevance', async () => {
      mockGraphClient.get.mockResolvedValue({
        value: [
          { id: '1', displayName: 'Work Notes' },
          { id: '2', displayName: 'Personal Work' },
          { id: '3', displayName: 'Random Stuff' }
        ]
      });

      const result = await handleSearchNotebooks('Work');
      const data = JSON.parse(result.content[0].text);
      
      expect(data.query).toBe('Work');
      expect(data.totalMatches).toBe(2);
      expect(data.results[0].displayName).toBe('Work Notes'); // Starts with "Work"
      expect(data.results[0]._matchScore).toBeGreaterThan(data.results[1]._matchScore);
    });

    it('returns exact matches with highest score', async () => {
      mockGraphClient.get.mockResolvedValue({
        value: [
          { id: '1', displayName: 'Work' },
          { id: '2', displayName: 'Work Notes' },
          { id: '3', displayName: 'My Work' }
        ]
      });

      const result = await handleSearchNotebooks('Work');
      const data = JSON.parse(result.content[0].text);
      
      expect(data.results[0].displayName).toBe('Work'); // Exact match
      expect(data.results[0]._matchScore).toBe(100);
    });

    it('respects the limit parameter', async () => {
      mockGraphClient.get.mockResolvedValue({
        value: [
          { id: '1', displayName: 'Test 1' },
          { id: '2', displayName: 'Test 2' },
          { id: '3', displayName: 'Test 3' }
        ]
      });

      const result = await handleSearchNotebooks('Test', 2);
      const data = JSON.parse(result.content[0].text);
      
      expect(data.totalMatches).toBe(3);
      expect(data.results.length).toBe(2);
    });
  });

  describe('handleSearchSections', () => {
    it('returns error when query is empty', async () => {
      const result = await handleSearchSections('');
      const data = JSON.parse(result.content[0].text);
      expect(data.error).toBe('query parameter is required');
    });

    it('searches all sections when no notebookId provided', async () => {
      mockGraphClient.get.mockResolvedValue({
        value: [
          { id: 's1', displayName: 'Meeting Notes' },
          { id: 's2', displayName: 'Project Plans' }
        ]
      });

      const result = await handleSearchSections('Notes');
      const data = JSON.parse(result.content[0].text);
      
      expect(data.notebookId).toBe('all');
      expect(data.totalMatches).toBe(1);
      expect(data.results[0].displayName).toBe('Meeting Notes');
    });

    it('searches within specific notebook when notebookId provided', async () => {
      mockGraphClient.get.mockResolvedValue({
        value: [
          { id: 's1', displayName: 'Chapter 1' }
        ]
      });

      const result = await handleSearchSections('Chapter', 'notebook123');
      const data = JSON.parse(result.content[0].text);
      
      expect(data.notebookId).toBe('notebook123');
      expect(mockGraphClient.api).toHaveBeenCalledWith('/me/onenote/notebooks/notebook123/sections');
    });
  });

  describe('handleSearchSectionGroups', () => {
    it('returns error when query is empty', async () => {
      const result = await handleSearchSectionGroups('');
      const data = JSON.parse(result.content[0].text);
      expect(data.error).toBe('query parameter is required');
    });

    it('returns matching section groups', async () => {
      mockGraphClient.get.mockResolvedValue({
        value: [
          { id: 'sg1', displayName: 'Archive 2024' },
          { id: 'sg2', displayName: 'Current Projects' }
        ]
      });

      const result = await handleSearchSectionGroups('Archive');
      const data = JSON.parse(result.content[0].text);
      
      expect(data.totalMatches).toBe(1);
      expect(data.results[0].displayName).toBe('Archive 2024');
    });
  });

  describe('handleUniversalSearch', () => {
    it('returns error when query is empty', async () => {
      const result = await handleUniversalSearch('');
      const data = JSON.parse(result.content[0].text);
      expect(data.error).toBe('query parameter is required');
    });

    it('searches all entity types by default', async () => {
      // Mock notebooks
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [{ id: 'n1', displayName: 'Project Alpha' }] }) // notebooks
        .mockResolvedValueOnce({ value: [{ id: 's1', displayName: 'Alpha Section' }] }) // sections
        .mockResolvedValueOnce({ value: [{ id: 'sg1', displayName: 'Alpha Group' }] }) // section groups
        .mockResolvedValueOnce({ value: [{ id: 'p1', title: 'Alpha Page' }] }); // pages

      const result = await handleUniversalSearch('Alpha');
      const data = JSON.parse(result.content[0].text);
      
      expect(data.searchedTypes).toEqual(['notebooks', 'sections', 'sectionGroups', 'pages']);
      expect(data.totalMatches).toBe(4);
      expect(data.groupedResults.notebooks.length).toBe(1);
      expect(data.groupedResults.sections.length).toBe(1);
      expect(data.groupedResults.sectionGroups.length).toBe(1);
      expect(data.groupedResults.pages.length).toBe(1);
    });

    it('filters by entity types when specified', async () => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [{ id: 'n1', displayName: 'My Notebook' }] }) // notebooks
        .mockResolvedValueOnce({ value: [{ id: 's1', displayName: 'My Section' }] }); // sections

      const result = await handleUniversalSearch('My', ['notebooks', 'sections']);
      const data = JSON.parse(result.content[0].text);
      
      expect(data.searchedTypes).toEqual(['notebooks', 'sections']);
      expect(data.groupedResults.notebooks.length).toBe(1);
      expect(data.groupedResults.sections.length).toBe(1);
      expect(data.groupedResults.sectionGroups.length).toBe(0);
      expect(data.groupedResults.pages.length).toBe(0);
    });

    it('respects the limit parameter', async () => {
      mockGraphClient.get
        .mockResolvedValueOnce({ 
          value: [
            { id: 'n1', displayName: 'Test Notebook 1' },
            { id: 'n2', displayName: 'Test Notebook 2' }
          ] 
        })
        .mockResolvedValueOnce({ 
          value: [
            { id: 's1', displayName: 'Test Section 1' },
            { id: 's2', displayName: 'Test Section 2' }
          ] 
        })
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({ value: [] });

      const result = await handleUniversalSearch('Test', undefined, undefined, 3);
      const data = JSON.parse(result.content[0].text);
      
      expect(data.totalMatches).toBe(4);
      expect(data.results.length).toBe(3);
    });
  });
});

describe('Match Score Calculation', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    const mockGraphClient = {
      api: vi.fn().mockReturnThis(),
      get: vi.fn()
    };
    vi.mocked(getGraphClient).mockReturnValue(mockGraphClient as any);
  });

  it('exact match gets score of 100', async () => {
    const mockGraphClient = {
      api: vi.fn().mockReturnThis(),
      get: vi.fn().mockResolvedValue({
        value: [{ id: '1', displayName: 'test' }]
      })
    };
    vi.mocked(getGraphClient).mockReturnValue(mockGraphClient as any);

    const result = await handleSearchNotebooks('test');
    const data = JSON.parse(result.content[0].text);
    expect(data.results[0]._matchScore).toBe(100);
  });

  it('starts-with match gets score of 80', async () => {
    const mockGraphClient = {
      api: vi.fn().mockReturnThis(),
      get: vi.fn().mockResolvedValue({
        value: [{ id: '1', displayName: 'testing stuff' }]
      })
    };
    vi.mocked(getGraphClient).mockReturnValue(mockGraphClient as any);

    const result = await handleSearchNotebooks('test');
    const data = JSON.parse(result.content[0].text);
    expect(data.results[0]._matchScore).toBe(80);
  });

  it('word boundary match gets score of 60', async () => {
    const mockGraphClient = {
      api: vi.fn().mockReturnThis(),
      get: vi.fn().mockResolvedValue({
        value: [{ id: '1', displayName: 'my test notebook' }]
      })
    };
    vi.mocked(getGraphClient).mockReturnValue(mockGraphClient as any);

    const result = await handleSearchNotebooks('test');
    const data = JSON.parse(result.content[0].text);
    expect(data.results[0]._matchScore).toBe(60);
  });

  it('contains match gets score of 40', async () => {
    const mockGraphClient = {
      api: vi.fn().mockReturnThis(),
      get: vi.fn().mockResolvedValue({
        value: [{ id: '1', displayName: 'mytesting' }]
      })
    };
    vi.mocked(getGraphClient).mockReturnValue(mockGraphClient as any);

    const result = await handleSearchNotebooks('test');
    const data = JSON.parse(result.content[0].text);
    expect(data.results[0]._matchScore).toBe(40);
  });

  it('case insensitive matching', async () => {
    const mockGraphClient = {
      api: vi.fn().mockReturnThis(),
      get: vi.fn().mockResolvedValue({
        value: [{ id: '1', displayName: 'TEST NOTEBOOK' }]
      })
    };
    vi.mocked(getGraphClient).mockReturnValue(mockGraphClient as any);

    const result = await handleSearchNotebooks('test');
    const data = JSON.parse(result.content[0].text);
    expect(data.results[0]._matchScore).toBe(80); // starts-with score
  });
});
