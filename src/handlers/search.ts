import { ensureGraphClient, getGraphClient } from '../lib/graph-client.js';
import { rateLimiter } from '../lib/rate-limiter.js';
import { successResponse, errorResponse, type ToolResponse } from './response.js';
import type { 
  OneNoteNotebook, 
  OneNoteSection, 
  OneNoteSectionGroup, 
  OneNotePage,
  OneNoteListResponse 
} from '../types.js';

/**
 * Search result with match information
 */
interface SearchMatch<T> {
  item: T;
  matchScore: number;
  matchedField: string;
}

/**
 * Calculate a simple match score for ranking results
 */
function calculateMatchScore(text: string | undefined, query: string): number {
  if (!text) return 0;
  const normalizedText = text.toLowerCase();
  const normalizedQuery = query.toLowerCase();
  
  // Exact match gets highest score
  if (normalizedText === normalizedQuery) return 100;
  
  // Starts with query gets high score
  if (normalizedText.startsWith(normalizedQuery)) return 80;
  
  // Contains query as whole word gets medium-high score
  const wordBoundary = new RegExp(`\\b${escapeRegex(normalizedQuery)}\\b`);
  if (wordBoundary.test(normalizedText)) return 60;
  
  // Contains query anywhere gets medium score
  if (normalizedText.includes(normalizedQuery)) return 40;
  
  return 0;
}

/**
 * Escape special regex characters
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Search for notebooks by name
 */
export async function handleSearchNotebooks(
  query: string,
  limit?: number
): Promise<ToolResponse> {
  try {
    if (!query || query.trim().length === 0) {
      console.log('[handleSearchNotebooks] Missing required parameter: query');
      return successResponse({ 
        error: 'query parameter is required',
        message: 'Please provide a search term to find notebooks.',
        example: 'searchNotebooks({ query: "work" })'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    console.log(`[handleSearchNotebooks] Searching for notebooks with query: "${query}"`);
    const response: OneNoteListResponse<OneNoteNotebook> = await rateLimiter.execute(() =>
      graphClient.api('/me/onenote/notebooks').get()
    );
    
    const notebooks = response.value || [];
    const searchTerm = query.toLowerCase().trim();
    
    // Score and filter notebooks
    const matches: SearchMatch<OneNoteNotebook>[] = [];
    for (const notebook of notebooks) {
      const score = calculateMatchScore(notebook.displayName, searchTerm);
      if (score > 0) {
        matches.push({
          item: notebook,
          matchScore: score,
          matchedField: 'displayName'
        });
      }
    }
    
    // Sort by score descending
    matches.sort((a, b) => b.matchScore - a.matchScore);
    
    // Apply limit if specified
    const results = limit && limit > 0 
      ? matches.slice(0, limit) 
      : matches;
    
    console.log(`[handleSearchNotebooks] Found ${matches.length} notebooks matching "${query}" (searched ${notebooks.length} total)`);
    return successResponse({
      query,
      totalMatches: matches.length,
      totalSearched: notebooks.length,
      results: results.map(m => ({
        ...m.item,
        _matchScore: m.matchScore,
        _matchedField: m.matchedField
      }))
    });
  } catch (error) {
    return errorResponse('Search notebooks', error, { query, limit, resourceType: 'notebook' });
  }
}

/**
 * Search for sections by name
 */
export async function handleSearchSections(
  query: string,
  notebookId?: string,
  limit?: number
): Promise<ToolResponse> {
  try {
    if (!query || query.trim().length === 0) {
      console.log('[handleSearchSections] Missing required parameter: query');
      return successResponse({ 
        error: 'query parameter is required',
        message: 'Please provide a search term to find sections.',
        example: 'searchSections({ query: "notes" })'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    let sections: OneNoteSection[] = [];
    const context = { query, notebookId, limit };
    console.log(`[handleSearchSections] Searching for sections with query: "${query}"`, context);
    
    if (notebookId) {
      // Search within a specific notebook
      const response: OneNoteListResponse<OneNoteSection> = await rateLimiter.execute(() =>
        graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).get()
      );
      sections = response.value || [];
    } else {
      // Search across all notebooks
      const response: OneNoteListResponse<OneNoteSection> = await rateLimiter.execute(() =>
        graphClient.api('/me/onenote/sections').get()
      );
      sections = response.value || [];
    }
    
    const searchTerm = query.toLowerCase().trim();
    
    // Score and filter sections
    const matches: SearchMatch<OneNoteSection>[] = [];
    for (const section of sections) {
      const score = calculateMatchScore(section.displayName, searchTerm);
      if (score > 0) {
        matches.push({
          item: section,
          matchScore: score,
          matchedField: 'displayName'
        });
      }
    }
    
    // Sort by score descending
    matches.sort((a, b) => b.matchScore - a.matchScore);
    
    // Apply limit if specified
    const results = limit && limit > 0 
      ? matches.slice(0, limit) 
      : matches;
    
    console.log(`[handleSearchSections] Found ${matches.length} sections matching "${query}" (searched ${sections.length} total)`);
    return successResponse({
      query,
      notebookId: notebookId || 'all',
      totalMatches: matches.length,
      totalSearched: sections.length,
      results: results.map(m => ({
        ...m.item,
        _matchScore: m.matchScore,
        _matchedField: m.matchedField
      }))
    });
  } catch (error) {
    return errorResponse('Search sections', error, { query, notebookId, limit, resourceType: 'section' });
  }
}

/**
 * Search for section groups by name
 */
export async function handleSearchSectionGroups(
  query: string,
  notebookId?: string,
  limit?: number
): Promise<ToolResponse> {
  try {
    if (!query || query.trim().length === 0) {
      console.log('[handleSearchSectionGroups] Missing required parameter: query');
      return successResponse({ 
        error: 'query parameter is required',
        message: 'Please provide a search term to find section groups.',
        example: 'searchSectionGroups({ query: "archive" })'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    let sectionGroups: OneNoteSectionGroup[] = [];
    const context = { query, notebookId, limit };
    console.log(`[handleSearchSectionGroups] Searching for section groups with query: "${query}"`, context);
    
    if (notebookId) {
      // Search within a specific notebook
      const response: OneNoteListResponse<OneNoteSectionGroup> = await rateLimiter.execute(() =>
        graphClient.api(`/me/onenote/notebooks/${notebookId}/sectionGroups`).get()
      );
      sectionGroups = response.value || [];
    } else {
      // Search across all notebooks
      const response: OneNoteListResponse<OneNoteSectionGroup> = await rateLimiter.execute(() =>
        graphClient.api('/me/onenote/sectionGroups').get()
      );
      sectionGroups = response.value || [];
    }
    
    const searchTerm = query.toLowerCase().trim();
    
    // Score and filter section groups
    const matches: SearchMatch<OneNoteSectionGroup>[] = [];
    for (const group of sectionGroups) {
      const score = calculateMatchScore(group.displayName, searchTerm);
      if (score > 0) {
        matches.push({
          item: group,
          matchScore: score,
          matchedField: 'displayName'
        });
      }
    }
    
    // Sort by score descending
    matches.sort((a, b) => b.matchScore - a.matchScore);
    
    // Apply limit if specified
    const results = limit && limit > 0 
      ? matches.slice(0, limit) 
      : matches;
    
    console.log(`[handleSearchSectionGroups] Found ${matches.length} section groups matching "${query}" (searched ${sectionGroups.length} total)`);
    return successResponse({
      query,
      notebookId: notebookId || 'all',
      totalMatches: matches.length,
      totalSearched: sectionGroups.length,
      results: results.map(m => ({
        ...m.item,
        _matchScore: m.matchScore,
        _matchedField: m.matchedField
      }))
    });
  } catch (error) {
    return errorResponse('Search section groups', error, { query, notebookId, limit, resourceType: 'sectionGroup' });
  }
}

/**
 * Universal search across all OneNote entities (notebooks, sections, section groups, pages)
 */
export async function handleUniversalSearch(
  query: string,
  entityTypes?: string[],
  notebookId?: string,
  limit?: number
): Promise<ToolResponse> {
  try {
    if (!query || query.trim().length === 0) {
      console.log('[handleUniversalSearch] Missing required parameter: query');
      return successResponse({ 
        error: 'query parameter is required',
        message: 'Please provide a search term to search across all OneNote entities.',
        example: 'universalSearch({ query: "project" })',
        availableTypes: ['notebooks', 'sections', 'sectionGroups', 'pages']
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    const searchTerm = query.toLowerCase().trim();
    const types = entityTypes || ['notebooks', 'sections', 'sectionGroups', 'pages'];
    const context = { query, entityTypes: types, notebookId, limit };
    console.log('[handleUniversalSearch] Starting universal search', {
      entityTypesCount: types.length,
      hasNotebookId: Boolean(notebookId),
      limit,
    });
    
    interface UniversalSearchResult {
      entityType: string;
      id: string;
      displayName: string;
      matchScore: number;
      parentInfo?: {
        notebookId?: string;
        notebookName?: string;
        sectionId?: string;
        sectionName?: string;
      };
      createdDateTime?: string;
      lastModifiedDateTime?: string;
    }
    
    const allResults: UniversalSearchResult[] = [];
    
    // Search notebooks
    if (types.includes('notebooks')) {
      try {
        const response: OneNoteListResponse<OneNoteNotebook> = await rateLimiter.execute(() =>
          graphClient.api('/me/onenote/notebooks').get()
        );
        
        for (const notebook of response.value || []) {
          const score = calculateMatchScore(notebook.displayName, searchTerm);
          if (score > 0) {
            allResults.push({
              entityType: 'notebook',
              id: notebook.id,
              displayName: notebook.displayName,
              matchScore: score,
              createdDateTime: notebook.createdDateTime,
              lastModifiedDateTime: notebook.lastModifiedDateTime
            });
          }
        }
      } catch (e) {
        console.warn(`[handleUniversalSearch] Failed to search notebooks: ${e instanceof Error ? e.message : 'Unknown error'}`);
      }
    }
    
    // Search sections
    if (types.includes('sections')) {
      try {
        const endpoint = notebookId 
          ? `/me/onenote/notebooks/${notebookId}/sections`
          : '/me/onenote/sections';
        
        const response: OneNoteListResponse<OneNoteSection> = await rateLimiter.execute(() =>
          graphClient.api(endpoint).get()
        );
        
        for (const section of response.value || []) {
          const score = calculateMatchScore(section.displayName, searchTerm);
          if (score > 0) {
            allResults.push({
              entityType: 'section',
              id: section.id,
              displayName: section.displayName,
              matchScore: score,
              parentInfo: section.parentNotebook ? {
                notebookId: section.parentNotebook.id,
                notebookName: section.parentNotebook.displayName
              } : undefined,
              createdDateTime: section.createdDateTime,
              lastModifiedDateTime: section.lastModifiedDateTime
            });
          }
        }
      } catch (e) {
        console.warn(`[handleUniversalSearch] Failed to search sections: ${e instanceof Error ? e.message : 'Unknown error'}`);
      }
    }
    
    // Search section groups
    if (types.includes('sectionGroups')) {
      try {
        const endpoint = notebookId
          ? `/me/onenote/notebooks/${notebookId}/sectionGroups`
          : '/me/onenote/sectionGroups';
        
        const response: OneNoteListResponse<OneNoteSectionGroup> = await rateLimiter.execute(() =>
          graphClient.api(endpoint).get()
        );
        
        for (const group of response.value || []) {
          const score = calculateMatchScore(group.displayName, searchTerm);
          if (score > 0) {
            allResults.push({
              entityType: 'sectionGroup',
              id: group.id,
              displayName: group.displayName,
              matchScore: score,
              createdDateTime: group.createdDateTime,
              lastModifiedDateTime: group.lastModifiedDateTime
            });
          }
        }
      } catch (e) {
        console.warn(`[handleUniversalSearch] Failed to search section groups: ${e instanceof Error ? e.message : 'Unknown error'}`);
      }
    }
    
    // Search pages
    if (types.includes('pages')) {
      try {
        let pages: OneNotePage[] = [];
        
        if (notebookId) {

          const sections = sectionsResponse.value || [];
          const concurrency = 5; // bounded concurrency to avoid too many simultaneous requests

          for (let i = 0; i < sections.length; i += concurrency) {
            const batch = sections.slice(i, i + concurrency);

            await Promise.all(
              batch.map(async (section) => {
                try {
                  const pagesResponse: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() =>
                    graphClient.api(`/me/onenote/sections/${section.id}/pages`).get()
                  );
                  for (const page of pagesResponse.value || []) {
                    page.sectionId = section.id;
                    page.sectionName = section.displayName;
                    pages.push(page);
                  }
                } catch (e) {
                  console.warn(
                    `[handleUniversalSearch] Failed to fetch pages from section ${section.displayName}: ${
                      e instanceof Error ? e.message : 'Unknown'
                    }`
                  );
                }
              })
            );
              }
            } catch (e) {
              console.warn(`[handleUniversalSearch] Failed to fetch pages from section ${section.displayName}: ${e instanceof Error ? e.message : 'Unknown'}`);
            }
          }
        } else {
          const response: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() =>
            graphClient.api('/me/onenote/pages').get()
          );
          pages = response.value || [];
        }
        
        for (const page of pages) {
          const score = calculateMatchScore(page.title, searchTerm);
          if (score > 0) {
            allResults.push({
              entityType: 'page',
              id: page.id,
              displayName: page.title,
              matchScore: score,
              parentInfo: {
                sectionId: page.sectionId || page.parentSection?.id,
                sectionName: page.sectionName || page.parentSection?.displayName
              },
              createdDateTime: page.createdDateTime,
              lastModifiedDateTime: page.lastModifiedDateTime
            });
          }
        }
      } catch (e) {
        console.warn(`[handleUniversalSearch] Failed to search pages: ${e instanceof Error ? e.message : 'Unknown error'}`);
      }
    }
    
    // Sort by score descending
    allResults.sort((a, b) => b.matchScore - a.matchScore);
    
    // Apply limit if specified
    const limitedResults = limit && limit > 0 
      ? allResults.slice(0, limit) 
      : allResults;
    
    // Group results by entity type for better readability
    const groupedResults = {
      notebooks: limitedResults.filter(r => r.entityType === 'notebook'),
      sections: limitedResults.filter(r => r.entityType === 'section'),
      sectionGroups: limitedResults.filter(r => r.entityType === 'sectionGroup'),
      pages: limitedResults.filter(r => r.entityType === 'page')
    };
    
    console.log(`[handleUniversalSearch] Found ${allResults.length} results matching "${query}" across ${types.join(', ')}`);
    return successResponse({
      query,
      searchedTypes: types,
      notebookId: notebookId || 'all',
      totalMatches: allResults.length,
      results: limitedResults,
      groupedResults
    });
  } catch (error) {
    return errorResponse('Universal search', error, { query, entityTypes, notebookId, limit, resourceType: 'mixed' });
  }
}
