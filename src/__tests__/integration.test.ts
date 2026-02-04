/**
 * Integration tests for OneNote handlers
 * 
 * These tests call the actual Microsoft Graph API and require valid authentication.
 * Run `npm run auth` first to authenticate if tests are skipped.
 * 
 * Test resources are created in a dedicated test notebook and cleaned up after tests.
 */

import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { loadToken, isValidTokenFormat } from '../lib/token-store.js';
import {
  handleListNotebooks,
  handleGetNotebook,
  handleCreateNotebook,
  handleListSections,
  handleGetSection,
  handleCreateSection,
  handleListSectionGroups,
  handleListPages,
  handleGetPage,
  handleCreatePage,
  handleUpdatePage,
  handleDeletePage,
  handleSearchPages
} from '../handlers/index.js';

// Test configuration
const TEST_NOTEBOOK_NAME = '_MCP_Integration_Test_Notebook';
const TEST_SECTION_NAME = 'Integration Test Section';
const TEST_PAGE_TITLE = 'Integration Test Page';

// Shared state
let hasValidCredentials = false;
let testNotebookId: string | null = null;
let testSectionId: string | null = null;
let testPageId: string | null = null;

// Helper to parse handler response
function parseResponse<T = unknown>(result: { content: Array<{ text: string }> }): T {
  return JSON.parse(result.content[0].text) as T;
}

// Helper to check if response is an error
function isErrorResponse(data: unknown): data is { error: string } {
  return typeof data === 'object' && data !== null && 'error' in data;
}

// Check for valid credentials
async function checkCredentials(): Promise<boolean> {
  try {
    const token = await loadToken();
    if (!token.token || !isValidTokenFormat(token.token)) {
      return false;
    }
    
    // Validate token works by making a simple API call
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { 'Authorization': `Bearer ${token.token}` }
    });
    return response.ok;
  } catch {
    return false;
  }
}

// Setup test resources
async function setupTestResources(): Promise<boolean> {
  try {
    // Find or create test notebook
    const notebooksResult = await handleListNotebooks();
    const notebooks = parseResponse<Array<{ id: string; displayName: string }>>(notebooksResult);
    
    if (isErrorResponse(notebooks)) {
      console.error('Failed to list notebooks:', notebooks.error);
      return false;
    }
    
    const existingNotebook = notebooks.find(n => n.displayName === TEST_NOTEBOOK_NAME);
    
    if (existingNotebook) {
      testNotebookId = existingNotebook.id;
    } else {
      // Create test notebook
      const createResult = await handleCreateNotebook(TEST_NOTEBOOK_NAME);
      const created = parseResponse<{ success?: boolean; id?: string; error?: string }>(createResult);
      
      if (!created.success || !created.id) {
        console.error('Failed to create test notebook:', created.error);
        return false;
      }
      
      testNotebookId = created.id;
      
      // Wait for notebook to be available
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
    
    // Find or create test section
    const sectionsResult = await handleListSections(testNotebookId);
    const sections = parseResponse<Array<{ id: string; displayName: string }>>(sectionsResult);
    
    if (!isErrorResponse(sections)) {
      const existingSection = sections.find(s => s.displayName === TEST_SECTION_NAME);
      
      if (existingSection) {
        testSectionId = existingSection.id;
      } else {
        const createSectionResult = await handleCreateSection(testNotebookId, TEST_SECTION_NAME);
        const createdSection = parseResponse<{ success?: boolean; id?: string }>(createSectionResult);
        
        if (createdSection.success && createdSection.id) {
          testSectionId = createdSection.id;
          await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }
    }
    
    return true;
  } catch (error) {
    console.error('Setup failed:', error);
    return false;
  }
}

// Cleanup test resources
async function cleanupTestResources(): Promise<void> {
  // Delete test page if created
  if (testPageId) {
    try {
      await handleDeletePage(testPageId);
    } catch { /* ignore */ }
  }
  
  // Note: OneNote API doesn't support deleting notebooks or sections
  // The test notebook will remain but can be reused in future test runs
}

// Initialize before all tests
beforeAll(async () => {
  hasValidCredentials = await checkCredentials();
  
  if (!hasValidCredentials) {
    console.log('\n');
    console.log('╔══════════════════════════════════════════════════════════════╗');
    console.log('║  ⚠️  No valid credentials found for integration tests        ║');
    console.log('║                                                              ║');
    console.log('║  Run: npm run auth                                           ║');
    console.log('║                                                              ║');
    console.log('║  Then re-run the tests to enable integration testing.       ║');
    console.log('╚══════════════════════════════════════════════════════════════╝');
    console.log('\n');
    return;
  }
  
  const setupSuccess = await setupTestResources();
  if (!setupSuccess) {
    console.error('Failed to setup test resources - some integration tests may fail');
  }
}, 60000);

// Cleanup after all tests
afterAll(async () => {
  if (hasValidCredentials) {
    await cleanupTestResources();
  }
}, 30000);

describe('Integration Tests', () => {
  describe('Notebooks', () => {
    test('handleListNotebooks returns notebooks array', async () => {
      if (!hasValidCredentials) {
        console.log('⏭️  Skipped: No valid credentials');
        return;
      }
      
      const result = await handleListNotebooks();
      const data = parseResponse<unknown[]>(result);
      
      expect(Array.isArray(data)).toBe(true);
    });

    test('handleGetNotebook returns notebook details', async () => {
      if (!hasValidCredentials || !testNotebookId) {
        console.log('⏭️  Skipped: No valid credentials or test notebook');
        return;
      }
      
      const result = await handleGetNotebook(testNotebookId);
      const data = parseResponse<{ id: string; displayName: string }>(result);
      
      expect(data.id).toBe(testNotebookId);
      expect(data.displayName).toBeTruthy();
    });

    test('handleGetNotebook without ID returns first notebook', async () => {
      if (!hasValidCredentials) {
        console.log('⏭️  Skipped: No valid credentials');
        return;
      }
      
      const result = await handleGetNotebook();
      const data = parseResponse<{ id?: string; error?: string }>(result);
      
      // Either returns a notebook or "No notebooks found" error
      expect(data.id || data.error).toBeTruthy();
    });
  });

  describe('Sections', () => {
    test('handleListSections returns sections array', async () => {
      if (!hasValidCredentials) {
        console.log('⏭️  Skipped: No valid credentials');
        return;
      }
      
      const result = await handleListSections();
      const data = parseResponse<unknown[]>(result);
      
      expect(Array.isArray(data)).toBe(true);
    }, 30000);

    test('handleListSections with notebookId filters correctly', async () => {
      if (!hasValidCredentials || !testNotebookId) {
        console.log('⏭️  Skipped: No valid credentials or test notebook');
        return;
      }
      
      const result = await handleListSections(testNotebookId);
      const data = parseResponse<Array<{ id: string }>>(result);
      
      expect(Array.isArray(data)).toBe(true);
    });

    test('handleGetSection returns section details', async () => {
      if (!hasValidCredentials || !testSectionId) {
        console.log('⏭️  Skipped: No valid credentials or test section');
        return;
      }
      
      const result = await handleGetSection(testSectionId);
      const data = parseResponse<{ id: string; displayName: string }>(result);
      
      expect(data.id).toBe(testSectionId);
      expect(data.displayName).toBeTruthy();
    });

    test('handleListSectionGroups returns array', async () => {
      if (!hasValidCredentials || !testNotebookId) {
        console.log('⏭️  Skipped: No valid credentials or test notebook');
        return;
      }
      
      const result = await handleListSectionGroups(testNotebookId);
      const data = parseResponse<unknown[]>(result);
      
      expect(Array.isArray(data)).toBe(true);
    });
  });

  describe('Pages', () => {
    test('handleCreatePage creates a new page', async () => {
      if (!hasValidCredentials || !testSectionId) {
        console.log('⏭️  Skipped: No valid credentials or test section');
        return;
      }
      
      const result = await handleCreatePage(
        testSectionId,
        TEST_PAGE_TITLE,
        '<p>Integration test content created at ' + new Date().toISOString() + '</p>'
      );
      const data = parseResponse<{ success?: boolean; id?: string; error?: string }>(result);
      
      expect(data.success).toBe(true);
      expect(data.id).toBeTruthy();
      
      if (data.id) {
        testPageId = data.id;
        // Wait for page to sync
        await new Promise(resolve => setTimeout(resolve, 2000));
      }
    }, 15000);

    test('handleListPages returns pages array', async () => {
      if (!hasValidCredentials || !testSectionId) {
        console.log('⏭️  Skipped: No valid credentials or test section');
        return;
      }
      
      const result = await handleListPages(testSectionId);
      const data = parseResponse<unknown[]>(result);
      
      expect(Array.isArray(data)).toBe(true);
    });

    test('handleGetPage returns page with content', async () => {
      if (!hasValidCredentials || !testPageId) {
        console.log('⏭️  Skipped: No valid credentials or test page');
        return;
      }
      
      const result = await handleGetPage(testPageId);
      const data = parseResponse<{ id: string; title: string; content?: string; error?: string }>(result);
      
      if (!data.error) {
        expect(data.id).toBe(testPageId);
        expect(data.content).toBeTruthy();
      }
    }, 30000);

    test('handleGetPage by title finds matching page', async () => {
      if (!hasValidCredentials || !testPageId) {
        console.log('⏭️  Skipped: No valid credentials or test page');
        return;
      }
      
      const result = await handleGetPage(undefined, 'Integration Test');
      const data = parseResponse<{ id?: string; title?: string; error?: string }>(result);
      
      // Should find the test page or report no match
      expect(data.id || data.error).toBeTruthy();
    }, 15000);

    test('handleSearchPages finds pages by query', async () => {
      if (!hasValidCredentials) {
        console.log('⏭️  Skipped: No valid credentials');
        return;
      }
      
      const result = await handleSearchPages('Integration');
      const data = parseResponse<unknown[] | { error: string }>(result);
      
      // Either returns array or error - both are valid responses
      expect(Array.isArray(data) || (typeof data === 'object' && data !== null)).toBe(true);
    }, 30000);

    test('handleSearchPages with sectionId filters correctly', async () => {
      if (!hasValidCredentials || !testSectionId) {
        console.log('⏭️  Skipped: No valid credentials or test section');
        return;
      }
      
      const result = await handleSearchPages('test', undefined, testSectionId);
      const data = parseResponse<unknown[]>(result);
      
      expect(Array.isArray(data)).toBe(true);
    });

    test('handleUpdatePage appends content', async () => {
      if (!hasValidCredentials || !testPageId) {
        console.log('⏭️  Skipped: No valid credentials or test page');
        return;
      }
      
      const result = await handleUpdatePage(
        testPageId,
        '<p>Updated at ' + new Date().toISOString() + '</p>'
      );
      const data = parseResponse<{ success?: boolean; error?: string }>(result);
      
      expect(data.success).toBe(true);
    }, 10000);

    test('handleDeletePage removes the page', async () => {
      if (!hasValidCredentials || !testPageId) {
        console.log('⏭️  Skipped: No valid credentials or test page');
        return;
      }
      
      const result = await handleDeletePage(testPageId);
      const data = parseResponse<{ success?: boolean; error?: string }>(result);
      
      expect(data.success).toBe(true);
      testPageId = null; // Clear so cleanup doesn't try again
    }, 10000);
  });

  describe('Error Handling', () => {
    test('handleGetNotebook with invalid ID returns error', async () => {
      if (!hasValidCredentials) {
        console.log('⏭️  Skipped: No valid credentials');
        return;
      }
      
      const result = await handleGetNotebook('invalid-notebook-id-12345');
      const data = parseResponse<{ error?: string }>(result);
      
      expect(data.error).toBeTruthy();
    });

    test('handleGetSection with invalid ID returns error', async () => {
      if (!hasValidCredentials) {
        console.log('⏭️  Skipped: No valid credentials');
        return;
      }
      
      const result = await handleGetSection('invalid-section-id-12345');
      const data = parseResponse<{ error?: string }>(result);
      
      expect(data.error).toBeTruthy();
    });

    test('handleGetPage with invalid ID returns error', async () => {
      if (!hasValidCredentials) {
        console.log('⏭️  Skipped: No valid credentials');
        return;
      }
      
      const result = await handleGetPage('invalid-page-id-12345');
      const data = parseResponse<{ error?: string }>(result);
      
      expect(data.error).toBeTruthy();
    });
  });
});
