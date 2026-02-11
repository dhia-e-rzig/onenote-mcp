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
    
    console.log(`[handleListSections] Fetching sections${notebookId ? ` for notebook: ${notebookId}` : ' (all)'}`);
    const response = await rateLimiter.execute(() => 
      graphClient.api(endpoint).get()
    );
    
    console.log(`[handleListSections] Successfully retrieved ${response.value?.length || 0} sections`);
    return successResponse(response.value);
  } catch (error) {
    return errorResponse('List sections', error, { notebookId, resourceType: 'section' });
  }
}

/**
 * Get a specific section by ID
 */
export async function handleGetSection(sectionId: string): Promise<ToolResponse> {
  try {
    if (!sectionId) {
      console.log('[handleGetSection] Missing required parameter: sectionId');
      return successResponse({ 
        error: 'sectionId is required',
        message: 'Please provide the ID of the section to retrieve.',
        suggestion: 'Use listSections to find available section IDs.'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    console.log(`[handleGetSection] Fetching section: ${sectionId}`);
    const response = await rateLimiter.execute(() => 
      graphClient.api(`/me/onenote/sections/${sectionId}`).get()
    );
    
    console.log(`[handleGetSection] Successfully retrieved section: ${response.displayName}`);
    return successResponse(response);
  } catch (error) {
    return errorResponse('Get section', error, { sectionId, resourceType: 'section' });
  }
}

/**
 * Create a new section in a notebook
 */
export async function handleCreateSection(notebookId: string, displayName: string): Promise<ToolResponse> {
  try {
    if (!notebookId) {
      console.log('[handleCreateSection] Missing required parameter: notebookId');
      return successResponse({ 
        error: 'notebookId is required',
        message: 'Please provide the ID of the notebook where you want to create the section.',
        suggestion: 'Use listNotebooks to find available notebook IDs.'
      });
    }
    if (!displayName) {
      console.log('[handleCreateSection] Missing required parameter: displayName');
      return successResponse({ 
        error: 'displayName is required',
        message: 'Please provide a name for the new section.',
        example: 'createSection({ notebookId: "...", displayName: "My Section" })'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    console.log(`[handleCreateSection] Creating section "${displayName}" in notebook: ${notebookId}`);
    const response = await rateLimiter.execute(() => 
      graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`).post({ displayName })
    );
    
    console.log(`[handleCreateSection] Successfully created section: ${response.displayName} (ID: ${response.id})`);
    return successResponse({
      success: true,
      id: response.id,
      displayName: response.displayName,
      self: response.self
    });
  } catch (error) {
    return errorResponse('Create section', error, { notebookId, displayName, resourceType: 'section' });
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
    
    console.log(`[handleListSectionGroups] Fetching section groups${notebookId ? ` for notebook: ${notebookId}` : ' (all)'}`);
    const response = await rateLimiter.execute(() => 
      graphClient.api(endpoint).get()
    );
    
    console.log(`[handleListSectionGroups] Successfully retrieved ${response.value?.length || 0} section groups`);
    return successResponse(response.value);
  } catch (error) {
    return errorResponse('List section groups', error, { notebookId, resourceType: 'sectionGroup' });
  }
}

/**
 * List all sections within a section group
 */
export async function handleListSectionsInGroup(sectionGroupId: string): Promise<ToolResponse> {
  try {
    if (!sectionGroupId) {
      console.log('[handleListSectionsInGroup] Missing required parameter: sectionGroupId');
      return successResponse({ 
        error: 'sectionGroupId is required',
        message: 'Please provide the ID of the section group.',
        suggestion: 'Use listSectionGroups to find available section group IDs.'
      });
    }
    
    await ensureGraphClient();
    const graphClient = getGraphClient()!;
    
    console.log(`[handleListSectionsInGroup] Fetching sections in group: ${sectionGroupId}`);
    const response = await rateLimiter.execute(() => 
      graphClient.api(`/me/onenote/sectionGroups/${sectionGroupId}/sections`).get()
    );
    
    console.log(`[handleListSectionsInGroup] Successfully retrieved ${response.value?.length || 0} sections`);
    return successResponse(response.value);
  } catch (error) {
    return errorResponse('List sections in group', error, { sectionGroupId, resourceType: 'section' });
  }
}
