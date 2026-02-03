# OneNote MCP Server

A Model Context Protocol (MCP) server implementation that enables AI language models like Claude and other LLMs to interact with Microsoft OneNote.

> This project is based on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks, with modifications to simplify authentication and improve usability.

## What Does This Do?

This server allows AI assistants to:
- Access your OneNote notebooks, sections, and pages
- Create new pages in your notebooks
- Search through your notes
- Read complete note content, including HTML formatting and text
- Analyze and summarize your notes directly

All of this happens directly through the AI interface without you having to switch contexts.

## Features

- **Automatic authentication** - Browser-based Microsoft login triggered automatically when needed
- **Token refresh** - Automatically refreshes tokens when expired, re-authenticates if refresh fails
- **No manual setup** - No Azure portal configuration or API keys required
- List all notebooks, sections, and pages
- Create new pages with HTML content
- Read complete page content, including HTML formatting
- Search across your notes

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

That's it! No additional configuration is needed.

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
      "args": ["Q:/Random Projects/onenote-mcp/onenote-mcp.mjs"]
    }
  }
}
```

> **Note:** Replace the path with the absolute path to your `onenote-mcp.mjs` file.

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
        "args": ["Q:/Random Projects/onenote-mcp/onenote-mcp.mjs"]
      }
    }
  }
}
```

## How Authentication Works

**Authentication is fully automatic.** When you ask the AI to interact with OneNote:

1. The MCP server checks for a valid token
2. If no token exists or it's expired, a browser window automatically opens
3. Sign in with your Microsoft account
4. The browser redirects back and authentication completes
5. Your request continues automatically

You don't need to run any authentication commands or manage tokens manually.

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

| Tool | Description |
|------|-------------|
| `listNotebooks` | List all your OneNote notebooks |
| `getNotebook` | Get details of a specific notebook |
| `listSections` | List all sections in a notebook |
| `listPages` | List all pages in a section |
| `getPage` | Get the complete content of a page (HTML) |
| `createPage` | Create a new page with HTML content |
| `searchPages` | Search for pages by title |

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

- This is normal - Microsoft tokens expire after ~1 hour
- The server automatically refreshes tokens when needed
- If refresh fails, you'll be prompted to sign in again

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
1. Go to [Azure Portal](https://portal.azure.com/) → Azure Active Directory → App registrations
2. Click "New registration"
3. Name your app (e.g., "OneNote MCP")
4. Select "Accounts in any organizational directory and personal Microsoft accounts"
5. Add redirect URI: `http://localhost:8400` (type: Web)
6. Copy the Application (client) ID
7. Under "API permissions", add:
   - `Notes.Read.All` (read operations)
   - `Notes.ReadWrite.All` (write operations)
   - `User.Read` (user profile)

## Direct Script Usage (Development)

For testing purposes, you can run standalone scripts:

```bash
# Manual authentication (creates token file)
node authenticate.js

# List notebooks
node simple-onenote.js
```

## Credits

Based on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks.

## License

MIT License - see LICENSE file for details.
