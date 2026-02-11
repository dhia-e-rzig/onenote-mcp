# OneNote MCP Server

A Model Context Protocol (MCP) server implementation that enables AI language models like Claude and other LLMs to interact with Microsoft OneNote.

> This project is TS reimplementation of  [onenote-mcp](https://github.com/danosb/onenote-mcp), with modifications to improve authentication and usability.

## What Does This Do?

This server allows AI assistants to:
- Access your OneNote notebooks, sections, and pages
- Create new pages in your notebooks
- Search through your notes
- Read complete note content, including HTML formatting and text
- Analyze and summarize your notes directly

All of this happens directly through the AI interface without you having to switch contexts.

## Features

- **TypeScript** - Fully typed codebase with strict type checking
- **Persistent authentication** - Authenticate once, stay logged in across sessions using refresh tokens
- **Secure token storage** - Tokens stored in OS credential manager (Windows Credential Manager, macOS Keychain, Linux Secret Service)
- **Automatic token refresh** - Silently refreshes expired tokens without user interaction
- **Browser-based login** - Simple Microsoft OAuth login when needed
- Full CRUD operations on notebooks, sections, and pages
- Search across your notes
- Rate limiting to prevent API throttling

## Installation

### Prerequisites

- Node.js 18 or higher (install from [nodejs.org](https://nodejs.org/))
- A Microsoft account with access to OneNote
- Git (install from [git-scm.com](https://git-scm.com/))

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/onenote-mcp.git
cd onenote-mcp
```

### Step 2: Install Dependencies

```bash
npm install
```

### Step 3: Build the Project

```bash
npm run build
```

### Step 4: Authenticate (One-Time Setup)

```bash
npm run auth
```

A browser window will open for Microsoft login. After signing in, your credentials are securely stored and will persist across sessions.

That's it! The server will automatically use your stored credentials.

## Setup for AI Assistants

### Cursor

1. Open Cursor Settings (Ctrl+, on Windows, Cmd+, on Mac)
2. Go to the **MCP** tab
3. Add a new MCP server with this configuration:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp/dist/index.js"]
    }
  }
}
```

> **Note:** Replace `/absolute/path/to/` with the actual path to your installation.

4. Restart Cursor

### Claude Desktop

1. Open Claude Desktop settings
2. Navigate to the MCP servers configuration
3. Add the OneNote server:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp/dist/index.js"]
    }
  }
}
```

4. Restart Claude Desktop

### VS Code with GitHub Copilot

Add to your VS Code `settings.json`:

```json
{
  "mcp": {
    "servers": {
      "onenote": {
        "command": "node",
        "args": ["/absolute/path/to/onenote-mcp/dist/index.js"]
      }
    }
  }
}
```

## How Authentication Works

**Authentication persists across sessions.** Run `npm run auth` once, and you're set:

1. Browser opens for Microsoft OAuth login
2. Sign in with your Microsoft account
3. Refresh token is securely stored in your OS credential manager
4. Future sessions automatically use the stored credentials

### Re-authentication

You only need to re-authenticate if:
- You revoke app permissions in your Microsoft account
- The refresh token expires (typically 90 days of inactivity)
- You manually clear credentials from your OS credential manager

### Token Management

- **Tokens are stored securely** in your operating system's credential manager:
  - Windows: Credential Manager
  - macOS: Keychain
  - Linux: Secret Service (libsecret)
- Tokens are automatically refreshed before they expire (5-minute buffer)
- If refresh fails, a new browser login is triggered
- Tokens are validated against Microsoft Graph API before use
- No plain-text files are used for token storage

## Available Tools

The following tools are available for AI assistants, organized by operation type:

### 📖 Read Operations

| Tool | Description | Parameters |
|------|-------------|------------|
| `listNotebooks` | List all your OneNote notebooks | - |
| `getNotebook` | Get details of a specific notebook | `notebookId` (optional) |
| `listSections` | List sections in a notebook | `notebookId` (optional) |
| `getSection` | Get details of a specific section | `sectionId` |
| `listSectionGroups` | List section groups in a notebook | `notebookId` (optional) |
| `listSectionsInGroup` | List sections within a section group | `sectionGroupId` |
| `listPages` | List pages in a section | `sectionId` (optional) |
| `getPage` | Get page content by ID or title | `pageId` or `title` |

### ✏️ Create Operations

| Tool | Description | Parameters |
|------|-------------|------------|
| `createNotebook` | Create a new notebook | `displayName` |
| `createSection` | Create a new section in a notebook | `notebookId`, `displayName` |
| `createPage` | Create a new page in a section | `sectionId`, `title` (optional), `content` (optional) |

### 🔄 Update Operations

| Tool | Description | Parameters |
|------|-------------|------------|
| `updatePage` | Append content to a page | `pageId`, `content`, `target` (optional) |

### 🗑️ Delete Operations

| Tool | Description | Parameters |
|------|-------------|------------|
| `deletePage` | Delete a page | `pageId` |

### 🔍 Search Operations

| Tool | Description | Parameters |
|------|-------------|------------|
| `search` | Universal search across all entity types | `query`, `entityTypes` (optional), `notebookId` (optional), `limit` (optional) |
| `searchNotebooks` | Search notebooks by name | `query`, `limit` (optional) |
| `searchSections` | Search sections by name | `query`, `notebookId` (optional), `limit` (optional) |
| `searchSectionGroups` | Search section groups by name | `query`, `notebookId` (optional), `limit` (optional) |
| `searchPages` | Search pages by title | `query`, `notebookId` (optional), `sectionId` (optional) |

#### Search Relevance Scoring

All search tools rank results by relevance:
- **100 points** - Exact match
- **80 points** - Name starts with search term
- **60 points** - Search term appears as a whole word
- **40 points** - Search term appears anywhere in the name

### 📋 Tool Lists by Permission Level

Use these lists to quickly configure which tools to enable based on access requirements.

#### Read Only (Read + Search)

```
"listNotebooks", "getNotebook", "listSections", "getSection", "listSectionGroups", "listSectionsInGroup", "listPages", "getPage", "search", "searchNotebooks", "searchSections", "searchSectionGroups", "searchPages"
```

#### Read/Write (All Operations)

```
"listNotebooks", "getNotebook", "listSections", "getSection", "listSectionGroups", "listSectionsInGroup", "listPages", "getPage", "createNotebook", "createSection", "createPage", "updatePage", "deletePage", "search", "searchNotebooks", "searchSections", "searchSectionGroups", "searchPages"
```

## Example Interactions

```
User: Show me my OneNote notebooks
AI: [uses listNotebooks] You have 3 notebooks: "Work", "Personal", and "Projects"

User: What's in my Work notebook?
AI: [uses listSections] Your Work notebook has these sections: "Meetings", "Tasks", "Notes"

User: Create a new page with today's meeting notes
AI: [uses createPage] Done! I've created a new page in your notebook.

User: Find my notes about the Q4 planning
AI: [uses searchPages] I found 2 pages matching "Q4 planning"...

User: Read the project requirements page and summarize it
AI: [uses getPage] Here's a summary of your project requirements: ...
```

## Development

### Project Structure

```
onenote-mcp/
├── src/
│   ├── index.ts              # MCP server entry point
│   ├── authenticate.ts       # OAuth authentication script
│   ├── types.ts              # TypeScript interfaces
│   ├── handlers/
│   │   ├── index.ts          # Handler exports
│   │   ├── response.ts       # Response helpers & types
│   │   ├── notebooks.ts      # Notebook operations
│   │   ├── sections.ts       # Section operations
│   │   ├── pages.ts          # Page operations
│   │   └── search.ts         # Search operations
│   ├── tools/
│   │   └── register.ts       # MCP tool registration
│   ├── lib/
│   │   ├── config.ts         # Configuration (MSAL, scopes)
│   │   ├── graph-client.ts   # Microsoft Graph client
│   │   ├── rate-limiter.ts   # API rate limiting
│   │   ├── token-store.ts    # Secure token persistence
│   │   └── validation.ts     # Input validation
│   └── __tests__/
│       ├── response.test.ts      # Response helper tests
│       ├── notebooks.test.ts     # Notebook handler tests
│       ├── sections.test.ts      # Section handler tests
│       ├── pages.test.ts         # Page handler tests
│       ├── search.test.ts        # Search handler tests
│       ├── token-store.test.ts   # Token store tests
│       └── integration.test.ts   # API integration tests
├── dist/                     # Compiled JavaScript (generated)
├── package.json
├── tsconfig.json
├── vitest.config.ts
└── eslint.config.mjs
```

### Scripts

```bash
npm run build         # Compile TypeScript to JavaScript
npm run dev           # Watch mode for development
npm run start         # Run the MCP server
npm run auth          # Authenticate with Microsoft
npm test              # Run all tests (unit + integration)
npm run test:watch    # Run tests in watch mode
npm run test:unit     # Run unit tests only (fast, mocked)
npm run test:integration  # Run integration tests (requires auth)
npm run lint          # Run ESLint
npm run lint:fix      # Run ESLint with auto-fix
npm run typecheck     # Type-check without emitting
npm run validate      # Run typecheck, lint, and all tests
```

### Running Tests

```bash
# Run all tests (75 unit + 18 integration)
npm test

# Run unit tests only (~1 second, no auth needed)
npm run test:unit

# Run integration tests only (~30 seconds, requires authentication)
npm run test:integration

# Watch mode for development
npm run test:watch
```

**Unit tests** mock external dependencies and test handler logic in isolation.

**Integration tests** call the real Microsoft Graph API. They require valid authentication (`npm run auth`) and will skip gracefully if credentials are not available.

## Troubleshooting

### Browser doesn't open for authentication

- Make sure you're running on a system with a GUI (not a headless server)
- Check that port 8400 is not in use by another application
- Try restarting the MCP server

### Authentication fails

- Clear your browser cookies for `login.microsoftonline.com`
- Make sure you're signing in with the correct Microsoft account
- Check your internet connection

### Token keeps expiring

- Refresh tokens last ~90 days of inactivity
- Run `npm run auth` to re-authenticate
- Check that your Microsoft account hasn't revoked app permissions

### Server won't start

- Verify Node.js 18+ is installed: `node --version`
- Build the project first: `npm run build`
- Reinstall dependencies: `rm -rf node_modules && npm install`
- Check for syntax errors in the console output

### AI can't connect to the server

- Ensure the path in your MCP configuration is an absolute path
- Make sure you're pointing to `dist/index.js` (the compiled output)
- Restart your AI assistant after configuration changes

## Security Notes

- **Tokens are stored in the OS credential manager** - not in plain text files
  - Windows: Windows Credential Manager
  - macOS: Keychain Access
  - Linux: Secret Service (requires libsecret)
- Tokens grant read/write access to your OneNote data
- Authentication uses Microsoft's standard OAuth 2.0 flow with PKCE
- No data is sent to third parties - only to Microsoft Graph API
- Input validation and rate limiting protect against abuse
- Error messages are sanitized to prevent information leakage
- To remove stored credentials:
  - Windows: Control Panel → Credential Manager → Windows Credentials → find "onenote-mcp"
  - macOS: Keychain Access → search "onenote-mcp"
  - Linux: Use `secret-tool` or your distribution's credential manager


## Credits

Based on [onenote-mcp](https://github.com/danosb/onenote-mcp) by danosb.

## License

MIT License - see LICENSE file for details.
