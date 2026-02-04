import { PublicClientApplication, CryptoProvider } from '@azure/msal-node';
import { saveToken } from './lib/token-store.js';
import http from 'http';
import open from 'open';

// Client ID - use environment variable for production
const clientId = process.env.AZURE_CLIENT_ID || '813d941f-92ac-4ac0-94a2-1e89b720e15b';

const scopes = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read'];
const redirectUri = 'http://localhost:8400';

// MSAL configuration
const msalConfig = {
  auth: {
    clientId: clientId,
    authority: 'https://login.microsoftonline.com/consumers'
  }
};

const pca = new PublicClientApplication(msalConfig);
const cryptoProvider = new CryptoProvider();

// Store PKCE codes for the auth flow
let pkceCodes = null;

async function authenticate() {
  console.log('Starting browser-based authentication...');
  console.log('A browser window will open for you to sign in with your Microsoft account.\n');
  
  return new Promise((resolve, reject) => {
    // Create a local server to handle the OAuth redirect
    const server = http.createServer(async (req, res) => {
      const url = new URL(req.url, redirectUri);
      
      if (url.pathname === '/' && url.searchParams.has('code')) {
        const code = url.searchParams.get('code');
        
        try {
          // Exchange authorization code for token
          const tokenRequest = {
            code: code,
            scopes: scopes,
            redirectUri: redirectUri,
            codeVerifier: pkceCodes.verifier
          };
          
          const response = await pca.acquireTokenByCode(tokenRequest);
          
          const expiresAt = response.expiresOn || new Date(Date.now() + 3600 * 1000);
          await saveToken(response.accessToken, expiresAt);
          
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Authentication Successful</title>
</head>
<body style="font-family: system-ui, -apple-system, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
  <div style="text-align: center; padding: 50px; background: white; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); max-width: 400px;">
    <div style="font-size: 64px; margin-bottom: 20px;">✅</div>
    <h1 style="color: #107c10; margin: 0 0 15px 0; font-size: 28px;">Authentication Successful!</h1>
    <p style="color: #666; margin: 0; font-size: 16px; line-height: 1.5;">You can close this window and return to your terminal.</p>
  </div>
</body>
</html>`);
          
          server.close();
          console.log('\nAuthentication successful!');
          console.log('Token stored securely in system credential manager.');
          resolve();
          
        } catch (error) {
          res.writeHead(500, { 'Content-Type': 'text/html' });
          res.end(`<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Authentication Failed</title>
</head>
<body style="font-family: system-ui, -apple-system, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);">
  <div style="text-align: center; padding: 50px; background: white; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); max-width: 400px;">
    <div style="font-size: 64px; margin-bottom: 20px;">❌</div>
    <h1 style="color: #d13438; margin: 0 0 15px 0; font-size: 28px;">Authentication Failed</h1>
    <p style="color: #666; margin: 0 0 10px 0; font-size: 16px;">Error: ${error.message}</p>
    <p style="color: #999; margin: 0; font-size: 14px;">Please close this window and try again.</p>
  </div>
</body>
</html>`);
          server.close();
          console.error('Failed to exchange code for token:', error.message);
          reject(error);
        }
        
      } else if (url.pathname === '/' && url.searchParams.has('error')) {
        const errorDescription = url.searchParams.get('error_description') || url.searchParams.get('error') || 'Unknown error';
        
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(`<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Authentication Failed</title>
</head>
<body style="font-family: system-ui, -apple-system, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);">
  <div style="text-align: center; padding: 50px; background: white; border-radius: 16px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); max-width: 400px;">
    <div style="font-size: 64px; margin-bottom: 20px;">❌</div>
    <h1 style="color: #d13438; margin: 0 0 15px 0; font-size: 28px;">Authentication Failed</h1>
    <p style="color: #666; margin: 0; font-size: 16px; line-height: 1.5;">${errorDescription}</p>
  </div>
</body>
</html>`);
        server.close();
        console.error('Authentication error:', errorDescription);
        reject(new Error(errorDescription));
        
      } else if (url.pathname === '/favicon.ico') {
        res.writeHead(204);
        res.end();
      }
    });
    
    server.listen(8400, async () => {
      console.log('Local server listening on http://localhost:8400');
      
      try {
        // Generate PKCE codes
        pkceCodes = await cryptoProvider.generatePkceCodes();
        
        const authCodeUrlParameters = {
          scopes: scopes,
          redirectUri: redirectUri,
          codeChallenge: pkceCodes.challenge,
          codeChallengeMethod: 'S256'
        };
        
        const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
        console.log('Opening browser for Microsoft login...');
        await open(authUrl);
        
      } catch (error) {
        server.close();
        console.error('Failed to generate auth URL:', error.message);
        reject(error);
      }
    });
    
    // Timeout after 2 minutes
    setTimeout(() => {
      server.close();
      reject(new Error('Authentication timed out after 2 minutes'));
    }, 120000);
  });
}

authenticate().catch(error => {
  console.error('Authentication failed:', error.message);
  process.exit(1);
}); 