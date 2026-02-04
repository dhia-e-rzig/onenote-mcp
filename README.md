# OneNote MCP Server

A Model Context Protocol (MCP) server implementation that enables AI language models like Claude and other LLMs to interact with Microsoft OneNote.

> This project is based on [onenote-mcp](https://github.com/danosb/onenote-mcp), with modifications to improve authentication and usability.

## What Does This Do?

This server allows AI assistants to:
- Access your OneNote notebooks, sections, and pages
- Create new pages in your notebooks
- Search through your notes
- Read complete note content, including HTML formatting and text
- Analyze and summarize your notes directly

All of this happens directly through the AI interface without you having to switch contexts.

## Features

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

### Step 3: Authenticate (One-Time Setup)

```bash
node authenticate.js
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
      "args": ["/absolute/path/to/onenote-mcp.mjs"]
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
      "args": ["/absolute/path/to/onenote-mcp.mjs"]
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
        "args": ["/absolute/path/to/onenote-mcp.mjs"]
      }
    }
  }
}
```

### GitHub Copilot CLI

Add to your `~/.config/github-copilot/config.json` (Linux/macOS) or `%APPDATA%\github-copilot\config.json` (Windows):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp.mjs"]
    }
  }
}
```

Then use with:
```bash
gh copilot chat --mcp onenote "List my OneNote notebooks"
```

### GitHub Copilot SDK (Node.js)

```javascript
import { CopilotClient } from '@github/copilot-sdk';
import { spawn } from 'child_process';

const client = new CopilotClient({
  mcpServers: {
    onenote: {
      command: 'node',
      args: ['/absolute/path/to/onenote-mcp.mjs']
    }
  }
});

// Use the OneNote tools
const notebooks = await client.callTool('onenote', 'listNotebooks', {});
console.log(notebooks);
```

### GitHub Copilot SDK (C#)

```csharp
using GitHub.Copilot.Sdk;

var client = new CopilotClientBuilder()
    .AddMcpServer("onenote", new McpServerConfig
    {
        Command = "node",
        Args = new[] { @"C:\path\to\onenote-mcp.mjs" }
    })
    .Build();

// Use the OneNote tools
var notebooks = await client.CallToolAsync("onenote", "listNotebooks", new { });
Console.WriteLine(notebooks);

// Create a page
var result = await client.CallToolAsync("onenote", "createPage", new 
{
    sectionId = "your-section-id",
    title = "Meeting Notes",
    content = "<p>Notes from today's meeting</p>"
});
```

## How Authentication Works

**Authentication persists across sessions.** Run `node authenticate.js` once, and you're set:

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

The following tools are available for AI assistants:

| Tool | Description | Parameters |
|------|-------------|------------|
| `listNotebooks` | List all your OneNote notebooks | - |
| `getNotebook` | Get details of a specific notebook | `notebookId` (optional) |
| `createNotebook` | Create a new notebook | `displayName` |
| `listSections` | List sections in a notebook | `notebookId` (optional) |
| `getSection` | Get details of a specific section | `sectionId` |
| `createSection` | Create a new section | `notebookId`, `displayName` |
| `listSectionGroups` | List section groups in a notebook | `notebookId` (optional) |
| `listPages` | List pages in a section | `sectionId` (optional) |
| `getPage` | Get page metadata | `pageId` or `title` |
| `getPageContent` | Get full page content (HTML) | `pageId` |
| `createPage` | Create a new page | `sectionId`, `title`, `content` (optional) |
| `updatePage` | Append content to a page | `pageId`, `content`, `target` (optional) |
| `deletePage` | Delete a page | `pageId` |
| `searchPages` | Search for pages by title | `query`, `notebookId` (optional), `sectionId` (optional) |

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
- Run `node authenticate.js` to re-authenticate
- Check that your Microsoft account hasn't revoked app permissions

### Server won't start

- Verify Node.js 18+ is installed: `node --version`
- Reinstall dependencies: `rm -rf node_modules && npm install`
- Check for syntax errors in the console output

### AI can't connect to the server

- Ensure the path in your MCP configuration is an absolute path
- Check that the file extension is `.mjs` not `.js`
- Restart your AI assistant after configuration changes

## Security Notes

- **Tokens are stored in the OS credential manager** - not in plain text files
  - Windows: Windows Credential Manager
  - macOS: Keychain Access
  - Linux: Secret Service (requires libsecret)
- Tokens grant read/write access to your OneNote data
- Authentication uses Microsoft's standard OAuth 2.0 flow
- No data is sent to third parties - only to Microsoft Graph API
- Input validation and rate limiting protect against abuse
- Error messages are sanitized to prevent information leakage
- To remove stored credentials:
  - Windows: Control Panel → Credential Manager → Windows Credentials → find "onenote-mcp"
  - macOS: Keychain Access → search "onenote-mcp"
  - Linux: Use `secret-tool` or your distribution's credential manager

### Client ID Configuration (Optional)

By default, this server uses Microsoft's Graph Explorer Client ID for authentication. For production use, you can register your own Azure AD application and set the client ID:

```bash
# Set your own Azure AD application client ID
export AZURE_CLIENT_ID="your-client-id-here"
```

Or in your MCP server configuration:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["path/to/onenote-mcp.mjs"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id-here"
      }
    }
  }
}
```

To register your own Azure AD application:
1. Go to [Azure Portal](https://portal.azure.com/) → Microsoft Entra ID → App registrations
2. Click "New registration"
3. Name your app (e.g., "OneNote MCP")
4. Select "Accounts in any organizational directory and personal Microsoft accounts"
5. Add redirect URI: `http://localhost:8400` (type: Mobile and desktop applications)
6. Copy the Application (client) ID
7. Under "API permissions", add Microsoft Graph permissions:
   - `Notes.ReadWrite` (read and write notebooks)
   - `User.Read` (user profile)
   - `offline_access` (refresh tokens)

## Direct Script Usage (Development)

For testing and development:

```bash
# Authenticate (one-time setup)
node authenticate.js

# Run the test suite
node test-api.js

# List notebooks (standalone test)
node simple-onenote.js
```

### Running Tests

The test suite validates authentication persistence and all MCP tools:

```bash
node test-api.js
```

Tests create a temporary `_MCP_Test_Notebook` that is reused across test runs. A link is provided at the end if you want to manually delete it.

## Configuration

All configuration is centralized in `lib/config.js`:

```javascript
// Azure AD settings
clientId        // Your app's client ID
authority       // Microsoft login endpoint
redirectUri     // OAuth redirect (http://localhost:8400)
scopes          // OAuth permissions

// Storage
keytarService   // Credential manager service name

// Rate limiting
rateLimitConfig // API throttling settings
```

## Credits

Based on [onenote-mcp](https://github.com/danosb/onenote-mcp) by danosb.

## License

MIT License - see LICENSE file for details.
