import { ensureGraphClient, getGraphClient } from '../lib/graph-client.js';
import { rateLimiter } from '../lib/rate-limiter.js';
import { successResponse, errorResponse, type ToolResponse } from './response.js';

/**
 * List all OneNote notebooks
 */
export async function handleListNotebooks(): Promise<ToolResponse> {
  try {
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    const response = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/notebooks').get()
    );
    return successResponse(response.value);
  } catch (error) {
    return errorResponse('List notebooks', error);
  }
}

/**
 * Get a specific notebook by ID, or the first notebook if no ID provided
 */
export async function handleGetNotebook(notebookId?: string): Promise<ToolResponse> {
  try {
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    if (notebookId) {
      const response = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/notebooks/${notebookId}`).get()
      );
      return successResponse(response);
    }
    
    const response = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/notebooks').get()
    );
    
    if (!response.value?.length) {
      return successResponse({ error: 'No notebooks found' });
    }
    
    return successResponse(response.value[0]);
  } catch (error) {
    return errorResponse('Get notebook', error);
  }
}

/**
 * Create a new notebook
 */
export async function handleCreateNotebook(displayName: string): Promise<ToolResponse> {
  try {
    if (!displayName) {
      return successResponse({ error: 'displayName is required' });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    const response = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/notebooks').post({ displayName })
    );
    
    return successResponse({
      success: true,
      id: response.id,
      displayName: response.displayName,
      self: response.self,
      links: response.links
    });
  } catch (error) {
    return errorResponse('Create notebook', error);
  }
}
