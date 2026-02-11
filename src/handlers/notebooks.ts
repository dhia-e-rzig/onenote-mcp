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
    console.log('[handleListNotebooks] Fetching all notebooks');
    const response = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/notebooks').get()
    );
    console.log(`[handleListNotebooks] Successfully retrieved ${response.value?.length || 0} notebooks`);
    return successResponse(response.value);
  } catch (error) {
    return errorResponse('List notebooks', error, { resourceType: 'notebook' });
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
      console.log(`[handleGetNotebook] Fetching notebook with ID: ${notebookId}`);
      const response = await rateLimiter.execute(() => 
        graphClient.api(`/me/onenote/notebooks/${notebookId}`).get()
      );
      console.log(`[handleGetNotebook] Successfully retrieved notebook: ${response.displayName}`);
      return successResponse(response);
    }
    
    console.log('[handleGetNotebook] No ID provided, fetching first available notebook');
    const response = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/notebooks').get()
    );
    
    if (!response.value?.length) {
      console.log('[handleGetNotebook] No notebooks found in account');
      return successResponse({ 
        error: 'No notebooks found', 
        message: 'Your OneNote account has no notebooks. Create one first using createNotebook.' 
      });
    }
    
    console.log(`[handleGetNotebook] Returning first notebook: ${response.value[0].displayName}`);
    return successResponse(response.value[0]);
  } catch (error) {
    return errorResponse('Get notebook', error, { notebookId, resourceType: 'notebook' });
  }
}

/**
 * Create a new notebook
 */
export async function handleCreateNotebook(displayName: string): Promise<ToolResponse> {
  try {
    if (!displayName) {
      console.log('[handleCreateNotebook] Missing required parameter: displayName');
      return successResponse({ 
        error: 'displayName is required',
        message: 'Please provide a name for the new notebook.',
        example: 'createNotebook({ displayName: "My Notebook" })'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    console.log(`[handleCreateNotebook] Creating notebook: ${displayName}`);
    const response = await rateLimiter.execute(() => 
      graphClient.api('/me/onenote/notebooks').post({ displayName })
    );
    
    console.log(`[handleCreateNotebook] Successfully created notebook: ${response.displayName} (ID: ${response.id})`);
    return successResponse({
      success: true,
      id: response.id,
      displayName: response.displayName,
      self: response.self,
      links: response.links
    });
  } catch (error) {
    return errorResponse('Create notebook', error, { displayName, resourceType: 'notebook' });
  }
}
