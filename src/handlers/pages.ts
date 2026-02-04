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
      const response = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get()
      );
      return successResponse(response.value);
    }
    
    // List pages from all sections
    const sectionsResponse = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/sections').get()
    );
    
    if (!sectionsResponse.value?.length) {
      return successResponse([]);
    }
    
    const allPages: OneNotePage[] = [];
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
      } catch { /* skip sections with errors */ }
    }
    
    return successResponse(allPages);
  } catch (error) {
    return errorResponse('List pages', error);
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
      try {
        targetPage = await rateLimiter.execute(() => 
          graphClient.api(`/me/onenote/pages/${pageId}`).get()
        );
      } catch {
        return successResponse({ error: `Page not found with ID: ${pageId}` });
      }
    } else if (title) {
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
        return successResponse({ error: `No page found matching title: ${title}` });
      }
    } else {
      return successResponse({ error: 'Please provide either pageId or title parameter' });
    }
    
    // Get page content
    const content = await rateLimiter.execute(async () => {
      const res = await fetch(
        `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage!.id}/content`,
        { headers: { 'Authorization': `Bearer ${accessToken}` } }
      );
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res.text();
    });
    
    return successResponse({
      id: targetPage!.id,
      title: targetPage!.title,
      createdDateTime: targetPage!.createdDateTime,
      lastModifiedDateTime: targetPage!.lastModifiedDateTime,
      contentUrl: targetPage!.contentUrl,
      content
    });
  } catch (error) {
    return errorResponse('Get page', error);
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
      return successResponse({ error: 'sectionId is required. Use listSections to find section IDs.' });
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
    
    const response = await rateLimiter.execute(() => 
      graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(html)
    );
    
    return successResponse({
      success: true,
      id: response.id,
      title: response.title,
      createdDateTime: response.createdDateTime,
      self: response.self
    });
  } catch (error) {
    return errorResponse('Create page', error);
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
      return successResponse({ error: 'pageId is required' });
    }
    if (!content) {
      return successResponse({ error: 'content is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    const patchContent: PagePatchContent[] = [{
      target: target || 'body',
      action: 'append',
      content: sanitizeHtmlContent(content)
    }];
    
    await rateLimiter.execute(() => 
      graphClient
        .api(`/me/onenote/pages/${pageId}/content`)
        .header('Content-Type', 'application/json')
        .patch(patchContent)
    );
    
    return successResponse({ success: true, message: 'Page updated successfully' });
  } catch (error) {
    return errorResponse('Update page', error);
  }
}

/**
 * Delete a page by ID
 */
export async function handleDeletePage(pageId: string): Promise<ToolResponse> {
  try {
    if (!pageId) {
      return successResponse({ error: 'pageId is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    await rateLimiter.execute(() => 
      graphClient.api(`/me/onenote/pages/${pageId}`).delete()
    );
    
    return successResponse({ success: true, message: 'Page deleted successfully' });
  } catch (error) {
    return errorResponse('Delete page', error);
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
      return successResponse({ error: 'query parameter is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    let pages: OneNotePage[] = [];
    
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
        } catch { /* skip */ }
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
    
    return successResponse(filtered);
  } catch (error) {
    return errorResponse('Search pages', error);
  }
}
