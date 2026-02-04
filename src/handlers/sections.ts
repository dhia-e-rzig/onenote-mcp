import { ensureGraphClient, getGraphClient } from '../lib/graph-client.js';
import { rateLimiter } from '../lib/rate-limiter.js';
import { successResponse, errorResponse, type ToolResponse } from './response.js';

/**
 * List all sections, optionally filtered by notebook
 */
export async function handleListSections(notebookId?: string): Promise<ToolResponse> {
  try {
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    const endpoint = notebookId 
      ? `/me/onenote/notebooks/${notebookId}/sections`
      : '/me/onenote/sections';
    
    const response = await rateLimiter.execute(() => 
      graphClient.api(endpoint).get()
    );
    
    return successResponse(response.value);
  } catch (error) {
    return errorResponse('List sections', error);
  }
}

/**
 * Get a specific section by ID
 */
export async function handleGetSection(sectionId: string): Promise<ToolResponse> {
  try {
    if (!sectionId) {
      return successResponse({ error: 'sectionId is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    const response = await rateLimiter.execute(() => 
      graphClient.api(`/me/onenote/sections/${sectionId}`).get()
    );
    
    return successResponse(response);
  } catch (error) {
    return errorResponse('Get section', error);
  }
}

/**
 * Create a new section in a notebook
 */
export async function handleCreateSection(notebookId: string, displayName: string): Promise<ToolResponse> {
  try {
    if (!notebookId) {
      return successResponse({ error: 'notebookId is required. Use listNotebooks to find notebook IDs.' });
    }
    if (!displayName) {
      return successResponse({ error: 'displayName is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    const response = await rateLimiter.execute(() => 
      graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).post({ displayName })
    );
    
    return successResponse({
      success: true,
      id: response.id,
      displayName: response.displayName,
      self: response.self
    });
  } catch (error) {
    return errorResponse('Create section', error);
  }
}

/**
 * List all section groups, optionally filtered by notebook
 */
export async function handleListSectionGroups(notebookId?: string): Promise<ToolResponse> {
  try {
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    const endpoint = notebookId
      ? `/me/onenote/notebooks/${notebookId}/sectionGroups`
      : '/me/onenote/sectionGroups';
    
    const response = await rateLimiter.execute(() => 
      graphClient.api(endpoint).get()
    );
    
    return successResponse(response.value);
  } catch (error) {
    return errorResponse('List section groups', error);
  }
}
