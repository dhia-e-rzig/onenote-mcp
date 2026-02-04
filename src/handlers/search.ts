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
      return successResponse({ error: 'query parameter is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
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
    
    return successResponse({
      query,
      totalMatches: matches.length,
      results: results.map(m => ({
        ...m.item,
        _matchScore: m.matchScore,
        _matchedField: m.matchedField
      }))
    });
  } catch (error) {
    return errorResponse('Search notebooks', error);
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
      return successResponse({ error: 'query parameter is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    let sections: OneNoteSection[] = [];
    
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
    
    return successResponse({
      query,
      notebookId: notebookId || 'all',
      totalMatches: matches.length,
      results: results.map(m => ({
        ...m.item,
        _matchScore: m.matchScore,
        _matchedField: m.matchedField
      }))
    });
  } catch (error) {
    return errorResponse('Search sections', error);
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
      return successResponse({ error: 'query parameter is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    let sectionGroups: OneNoteSectionGroup[] = [];
    
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
    
    return successResponse({
      query,
      notebookId: notebookId || 'all',
      totalMatches: matches.length,
      results: results.map(m => ({
        ...m.item,
        _matchScore: m.matchScore,
        _matchedField: m.matchedField
      }))
    });
  } catch (error) {
    return errorResponse('Search section groups', error);
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
      return successResponse({ error: 'query parameter is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    const searchTerm = query.toLowerCase().trim();
    const types = entityTypes || ['notebooks', 'sections', 'sectionGroups', 'pages'];
    
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
      } catch { /* skip on error */ }
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
      } catch { /* skip on error */ }
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
      } catch { /* skip on error */ }
    }
    
    // Search pages
    if (types.includes('pages')) {
      try {
        let pages: OneNotePage[] = [];
        
        if (notebookId) {
          // Get sections in notebook first, then pages
          const sectionsResponse = await rateLimiter.execute(() =>
            graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).get()
          );
          
          for (const section of sectionsResponse.value || []) {
            try {
              const pagesResponse: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() =>
                graphClient.api(`/me/onenote/sections/${section.id}/pages`).get()
              );
              for (const page of pagesResponse.value || []) {
                page.sectionId = section.id;
                page.sectionName = section.displayName;
                pages.push(page);
              }
            } catch { /* skip */ }
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
      } catch { /* skip on error */ }
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
    
    return successResponse({
      query,
      searchedTypes: types,
      notebookId: notebookId || 'all',
      totalMatches: allResults.length,
      results: limitedResults,
      groupedResults
    });
  } catch (error) {
    return errorResponse('Universal search', error);
  }
}
