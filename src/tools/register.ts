import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import {
  handleListNotebooks,
  handleGetNotebook,
  handleCreateNotebook,
  handleListSections,
  handleGetSection,
  handleCreateSection,
  handleListSectionGroups,
  handleListSectionsInGroup,
  handleListPages,
  handleGetPage,
  handleCreatePage,
  handleUpdatePage,
  handleDeletePage,
  handleSearchPages,
  handleSearchNotebooks,
  handleSearchSections,
  handleSearchSectionGroups,
  handleUniversalSearch
} from '../handlers/index.js';

/**
 * Register all OneNote tools with the MCP server
 */
export function registerTools(server: McpServer): void {
  // ============================================
  // Notebook Tools
  // ============================================

  server.registerTool(
    'listNotebooks',
    { description: 'List all OneNote notebooks' },
    handleListNotebooks
  );

  server.registerTool(
    'getNotebook',
    {
      description: 'Get details of a specific notebook by ID',
      inputSchema: {
        notebookId: z.string().optional().describe('The ID of the notebook to get. If not provided, returns the first notebook.')
      }
    },
    ({ notebookId }) => handleGetNotebook(notebookId)
  );

  server.registerTool(
    'createNotebook',
    {
      description: 'Create a new notebook',
      inputSchema: {
        displayName: z.string().describe('The name of the new notebook. Required.')
      }
    },
    ({ displayName }) => handleCreateNotebook(displayName)
  );

  // ============================================
  // Section Tools
  // ============================================

  server.registerTool(
    'listSections',
    {
      description: 'List all sections. Provide a notebookId to list sections from a specific notebook, or omit to list all sections.',
      inputSchema: {
        notebookId: z.string().optional().describe('The ID of the notebook to list sections from. If not provided, lists all sections across all notebooks.')
      }
    },
    ({ notebookId }) => handleListSections(notebookId)
  );

  server.registerTool(
    'getSection',
    {
      description: 'Get details of a specific section by ID',
      inputSchema: {
        sectionId: z.string().describe('The ID of the section to get. Required.')
      }
    },
    ({ sectionId }) => handleGetSection(sectionId)
  );

  server.registerTool(
    'createSection',
    {
      description: 'Create a new section in a notebook',
      inputSchema: {
        notebookId: z.string().describe('The ID of the notebook to create the section in. Required.'),
        displayName: z.string().describe('The name of the new section. Required.')
      }
    },
    ({ notebookId, displayName }) => handleCreateSection(notebookId, displayName)
  );

  server.registerTool(
    'listSectionGroups',
    {
      description: 'List all section groups. Provide a notebookId to list section groups from a specific notebook.',
      inputSchema: {
        notebookId: z.string().optional().describe('The ID of the notebook to list section groups from. If not provided, lists all section groups.')
      }
    },
    ({ notebookId }) => handleListSectionGroups(notebookId)
  );

  server.registerTool(
    'listSectionsInGroup',
    {
      description: 'List all sections within a specific section group. Section groups are folders that can contain sections.',
      inputSchema: {
        sectionGroupId: z.string().describe('The ID of the section group to list sections from. Required. Use listSectionGroups to find section group IDs.')
      }
    },
    ({ sectionGroupId }) => handleListSectionsInGroup(sectionGroupId)
  );

  // ============================================
  // Page Tools
  // ============================================

  server.registerTool(
    'listPages',
    {
      description: 'List all pages in a section. Provide a sectionId to list pages from a specific section, or omit to list pages from all sections.',
      inputSchema: {
        sectionId: z.string().optional().describe('The ID of the section to list pages from. If not provided, lists pages from all sections.')
      }
    },
    ({ sectionId }) => handleListPages(sectionId)
  );

  server.registerTool(
    'getPage',
    {
      description: 'Get the content of a page by ID or title',
      inputSchema: {
        pageId: z.string().optional().describe('The ID of the page to retrieve. Takes precedence over title search.'),
        title: z.string().optional().describe('Search for a page by title (partial match). Used if pageId is not provided.')
      }
    },
    ({ pageId, title }) => handleGetPage(pageId, title)
  );

  server.registerTool(
    'createPage',
    {
      description: 'Create a new page in a specific section',
      inputSchema: {
        sectionId: z.string().describe('The ID of the section to create the page in. Required.'),
        title: z.string().optional().describe('The title of the new page. Defaults to "New Page" if not provided.'),
        content: z.string().optional().describe('The HTML content for the page body. Can include basic HTML tags like <p>, <h1>, <ul>, etc.')
      }
    },
    ({ sectionId, title, content }) => handleCreatePage(sectionId, title, content)
  );

  server.registerTool(
    'updatePage',
    {
      description: 'Update a page by appending content to it',
      inputSchema: {
        pageId: z.string().describe('The ID of the page to update. Required.'),
        content: z.string().describe('The HTML content to append to the page. Required.'),
        target: z.string().optional().describe('Where to insert the content: "body" (default) appends to page body, or specify an element ID.')
      }
    },
    ({ pageId, content, target }) => handleUpdatePage(pageId, content, target)
  );

  server.registerTool(
    'deletePage',
    {
      description: 'Delete a page by ID',
      inputSchema: {
        pageId: z.string().describe('The ID of the page to delete. Required.')
      }
    },
    ({ pageId }) => handleDeletePage(pageId)
  );

  server.registerTool(
    'searchPages',
    {
      description: 'Search for pages by title across all notebooks',
      inputSchema: {
        query: z.string().describe('The search term to find in page titles. Required.'),
        notebookId: z.string().optional().describe('Optional: Limit search to pages within a specific notebook.'),
        sectionId: z.string().optional().describe('Optional: Limit search to pages within a specific section.')
      }
    },
    ({ query, notebookId, sectionId }) => handleSearchPages(query, notebookId, sectionId)
  );

  // ============================================
  // Search Tools
  // ============================================

  server.registerTool(
    'searchNotebooks',
    {
      description: 'Search for notebooks by name. Returns notebooks matching the search query, ranked by relevance.',
      inputSchema: {
        query: z.string().describe('The search term to find in notebook names. Required.'),
        limit: z.number().optional().describe('Optional: Maximum number of results to return.')
      }
    },
    ({ query, limit }) => handleSearchNotebooks(query, limit)
  );

  server.registerTool(
    'searchSections',
    {
      description: 'Search for sections by name. Returns sections matching the search query, ranked by relevance.',
      inputSchema: {
        query: z.string().describe('The search term to find in section names. Required.'),
        notebookId: z.string().optional().describe('Optional: Limit search to sections within a specific notebook.'),
        limit: z.number().optional().describe('Optional: Maximum number of results to return.')
      }
    },
    ({ query, notebookId, limit }) => handleSearchSections(query, notebookId, limit)
  );

  server.registerTool(
    'searchSectionGroups',
    {
      description: 'Search for section groups by name. Returns section groups matching the search query, ranked by relevance.',
      inputSchema: {
        query: z.string().describe('The search term to find in section group names. Required.'),
        notebookId: z.string().optional().describe('Optional: Limit search to section groups within a specific notebook.'),
        limit: z.number().optional().describe('Optional: Maximum number of results to return.')
      }
    },
    ({ query, notebookId, limit }) => handleSearchSectionGroups(query, notebookId, limit)
  );

  server.registerTool(
    'search',
    {
      description: 'Universal search across all OneNote entities (notebooks, sections, section groups, and pages). Returns all matching items ranked by relevance. Use this when you want to find something but are unsure of its type.',
      inputSchema: {
        query: z.string().describe('The search term to find across all entities. Required.'),
        entityTypes: z.array(z.enum(['notebooks', 'sections', 'sectionGroups', 'pages'])).optional()
          .describe('Optional: Array of entity types to search. Defaults to all types. Options: "notebooks", "sections", "sectionGroups", "pages".'),
        notebookId: z.string().optional().describe('Optional: Limit search to within a specific notebook.'),
        limit: z.number().optional().describe('Optional: Maximum number of results to return.')
      }
    },
    ({ query, entityTypes, notebookId, limit }) => handleUniversalSearch(query, entityTypes, notebookId, limit)
  );
}
