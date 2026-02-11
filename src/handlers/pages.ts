import { ensureGraphClient, getGraphClient, getAccessToken } from '../lib/graph-client.js';
import { rateLimiter } from '../lib/rate-limiter.js';
import { sanitizeHtmlContent } from '../lib/validation.js';
import { successResponse, errorResponse, type ToolResponse } from './response.js';
import type { OneNotePage, OneNoteListResponse, PagePatchContent } from '../types.js';

/**
 * List all pages, optionally filtered by section
 */
export async function handleListPages(sectionId?: string): Promise<ToolResponse> {
  try {
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    if (sectionId) {
      console.log(`[handleListPages] Fetching pages for section: ${sectionId}`);
      const response = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get()
      );
      console.log(`[handleListPages] Successfully retrieved ${response.value?.length || 0} pages`);
      return successResponse(response.value);
    }
    
    // List pages from all sections
    console.log('[handleListPages] Fetching pages from all sections');
    const sectionsResponse = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/sections').get()
    );
    
    if (!sectionsResponse.value?.length) {
      console.log('[handleListPages] No sections found, returning empty list');
      return successResponse([]);
    }
    
    const allPages: OneNotePage[] = [];
    let sectionsProcessed = 0;
    let sectionsWithErrors = 0;
    for (const section of sectionsResponse.value) {
      try {
        const pagesResponse: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/sections/${section.id}/pages`).get()
        );
        if (pagesResponse.value) {
          for (const page of pagesResponse.value) {
            page.sectionName = section.displayName;
            page.sectionId = section.id;
            allPages.push(page);
          }
        }
        sectionsProcessed++;
      } catch (e) {
        sectionsWithErrors++;
        console.warn(`[handleListPages] Failed to fetch pages from section ${section.displayName}: ${e instanceof Error ? e.message : 'Unknown error'}`);
      }
    }
    
    console.log(`[handleListPages] Successfully retrieved ${allPages.length} pages from ${sectionsProcessed} sections (${sectionsWithErrors} sections had errors)`);
    return successResponse(allPages);
  } catch (error) {
    return errorResponse('List pages', error, { sectionId, resourceType: 'page' });
  }
}

/**
 * Get a page by ID or title search
 */
export async function handleGetPage(pageId?: string, title?: string): Promise<ToolResponse> {
  try {
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    const accessToken = getAccessToken();
    
    let targetPage: OneNotePage | null = null;
    
    if (pageId) {
      console.log(`[handleGetPage] Fetching page by ID: ${pageId}`);
      try {
        targetPage = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/pages/${pageId}`).get()
        );
      } catch (e) {
        console.log(`[handleGetPage] Page not found with ID: ${pageId}, error: ${e instanceof Error ? e.message : 'Unknown'}`);
        return successResponse({ 
          error: `Page not found with ID: ${pageId}`,
          message: 'The page may have been deleted or the ID is incorrect.',
          suggestion: 'Use listPages or searchPages to find valid page IDs.'
        });
      }
    } else if (title) {
      console.log(`[handleGetPage] Searching for page by title: ${title}`);
      const pagesResponse: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() => 
        graphClient.api('/me/onenote/pages').get()
      );
      
      if (pagesResponse.value?.length) {
        const searchLower = title.toLowerCase();
        targetPage = pagesResponse.value.find(p => 
          p.title?.toLowerCase().includes(searchLower)
        ) || null;
      }
      
      if (!targetPage) {
        console.log(`[handleGetPage] No page found matching title: ${title}`);
        return successResponse({ 
          error: `No page found matching title: ${title}`,
          message: 'No pages were found that match your search term.',
          suggestion: 'Try a different search term or use listPages to see all available pages.',
          searchedTitle: title
        });
      }
    } else {
      console.log('[handleGetPage] No pageId or title provided');
      return successResponse({ 
        error: 'Please provide either pageId or title parameter',
        message: 'You must specify which page to retrieve.',
        suggestion: 'Use pageId for exact lookup, or title for a title search.'
      });
    }
    
    // Get page content
    console.log(`[handleGetPage] Fetching content for page: ${targetPage!.title}`);
    const content = await rateLimiter.execute(async () => {
      const res = await fetch(
        `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage!.id}/content`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );
      if (!res.ok) {
        throw new Error(`HTTP ${res.status}: Failed to fetch page content`);
      }
      return res.text();
    });
    
    console.log(`[handleGetPage] Successfully retrieved page: ${targetPage!.title}`);
    return successResponse({
      id: targetPage!.id,
      title: targetPage!.title,
      createdDateTime: targetPage!.createdDateTime,
      lastModifiedDateTime: targetPage!.lastModifiedDateTime,
      contentUrl: targetPage!.contentUrl,
      content
    });
  } catch (error) {
    return errorResponse('Get page', error, { pageId, title, resourceType: 'page' });
  }
}

/**
 * Create a new page in a section
 */
export async function handleCreatePage(
  sectionId: string,
  title?: string,
  content?: string
): Promise<ToolResponse> {
  try {
    if (!sectionId) {
      console.log('[handleCreatePage] Missing required parameter: sectionId');
      return successResponse({ 
        error: 'sectionId is required',
        message: 'Please provide the ID of the section where you want to create the page.',
        suggestion: 'Use listSections to find available section IDs.'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    const pageTitle = title || 'New Page';
    const bodyContent = content || '<p>This is a new page created via the OneNote MCP.</p>';
    
    const html = sanitizeHtmlContent(`
      <!DOCTYPE html>
      <html>
        <head><title>${pageTitle}</title></head>
        <body>${bodyContent}</body>
      </html>
    `);
    
    console.log(`[handleCreatePage] Creating page "${pageTitle}" in section: ${sectionId}`);
    const response = await rateLimiter.execute(() => 
      graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(html)
    );
    
    console.log(`[handleCreatePage] Successfully created page: ${response.title} (ID: ${response.id})`);
    return successResponse({
      success: true,
      id: response.id,
      title: response.title,
      createdDateTime: response.createdDateTime,
      self: response.self
    });
  } catch (error) {
    return errorResponse('Create page', error, { sectionId, title, resourceType: 'page' });
  }
}

/**
 * Update a page by appending content
 */
export async function handleUpdatePage(
  pageId: string,
  content: string,
  target?: string
): Promise<ToolResponse> {
  try {
    if (!pageId) {
      console.log('[handleUpdatePage] Missing required parameter: pageId');
      return successResponse({ 
        error: 'pageId is required',
        message: 'Please provide the ID of the page to update.',
        suggestion: 'Use listPages or searchPages to find page IDs.'
      });
    }
    if (!content) {
      console.log('[handleUpdatePage] Missing required parameter: content');
      return successResponse({ 
        error: 'content is required',
        message: 'Please provide the HTML content to append to the page.',
        example: 'updatePage({ pageId: "...", content: "<p>New content</p>" })'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    const patchContent: PagePatchContent[] = [{
      target: target || 'body',
      action: 'append',
      content: sanitizeHtmlContent(content)
    }];
    
    console.log(`[handleUpdatePage] Updating page: ${pageId} (target: ${target || 'body'})`);
    await rateLimiter.execute(() => 
      graphClient
        .api(`/me/onenote/pages/${pageId}/content`)
        .header('Content-Type', 'application/json')
        .patch(patchContent)
    );
    
    console.log(`[handleUpdatePage] Successfully updated page: ${pageId}`);
    return successResponse({ success: true, message: 'Page updated successfully', pageId });
  } catch (error) {
    return errorResponse('Update page', error, { pageId, target, resourceType: 'page' });
  }
}

/**
 * Delete a page by ID
 */
export async function handleDeletePage(pageId: string): Promise<ToolResponse> {
  try {
    if (!pageId) {
      console.log('[handleDeletePage] Missing required parameter: pageId');
      return successResponse({ 
        error: 'pageId is required',
        message: 'Please provide the ID of the page to delete.',
        suggestion: 'Use listPages or searchPages to find page IDs.'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    console.log(`[handleDeletePage] Deleting page: ${pageId}`);
    await rateLimiter.execute(() => 
      graphClient.api(`/me/onenote/pages/${pageId}`).delete()
    );
    
    console.log(`[handleDeletePage] Successfully deleted page: ${pageId}`);
    return successResponse({ success: true, message: 'Page deleted successfully', deletedPageId: pageId });
  } catch (error) {
    return errorResponse('Delete page', error, { pageId, resourceType: 'page' });
  }
}

/**
 * Search for pages by title
 */
export async function handleSearchPages(
  query: string,
  notebookId?: string,
  sectionId?: string
): Promise<ToolResponse> {
  try {
    if (!query) {
      console.log('[handleSearchPages] Missing required parameter: query');
      return successResponse({ 
        error: 'query parameter is required',
        message: 'Please provide a search term to find pages.',
        example: 'searchPages({ query: "meeting notes" })'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    let pages: OneNotePage[] = [];
    const context = { query, notebookId, sectionId };
    console.log(`[handleSearchPages] Searching for pages with query: "${query}"`, context);
    
    if (sectionId) {
      const response: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get()
      );
      pages = response.value || [];
    } else if (notebookId) {
      const sectionsResponse = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).get()
      );
      
      for (const section of sectionsResponse.value || []) {
        try {
          const pagesResponse: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() => 
            graphClient.api(`/me/onenote/sections/${section.id}/pages`).get()
          );
          if (pagesResponse.value) {
            for (const page of pagesResponse.value) {
              page.sectionName = section.displayName;
              page.sectionId = section.id;
              pages.push(page);
            }
          }
        } catch (e) {
          console.warn(`[handleSearchPages] Failed to fetch pages from section ${section.displayName}: ${e instanceof Error ? e.message : 'Unknown'}`);
        }
      }
    } else {
      const response: OneNoteListResponse<OneNotePage> = await rateLimiter.execute(() => 
        graphClient.api('/me/onenote/pages').get()
      );
      pages = response.value || [];
    }
    
    const searchTerm = query.toLowerCase();
    const filtered = pages.filter(page => 
      page.title?.toLowerCase().includes(searchTerm)
    );
    
    console.log(`[handleSearchPages] Found ${filtered.length} pages matching "${query}" (searched ${pages.length} total)`);
    return successResponse({
      query,
      matches: filtered.length,
      totalSearched: pages.length,
      results: filtered
    });
  } catch (error) {
    return errorResponse('Search pages', error, { query, notebookId, sectionId, resourceType: 'page' });
  }
}
