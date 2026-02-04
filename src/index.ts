#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { registerTools } from './tools/register.js';

// Load environment variables
dotenv.config();

// Warn if using default client ID
if (!process.env.AZURE_CLIENT_ID) {
  console.error('WARNING: Using default Microsoft Graph Explorer client ID.');
  console.error('For production, register your own Azure AD application and set AZURE_CLIENT_ID environment variable.');
}

// Create MCP Server
const server = new McpServer(
  { name: 'onenote', version: '1.0.0', description: 'OneNote MCP Server' },
  { capabilities: { tools: { listChanged: true } } }
);

// Register all tools
registerTools(server);

// Entry point
async function main(): Promise<void> {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('OneNote MCP Server started successfully.');
    console.error('Authentication will be triggered automatically when needed.');
    
    process.on('SIGINT', () => process.exit(0));
  } catch {
    console.error('Server failed to start');
    process.exit(1);
  }
}

main();

