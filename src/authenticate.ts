#!/usr/bin/env node

import { PublicClientApplication, CryptoProvider, type AuthenticationResult } from '@azure/msal-node';
import { saveToken, saveRefreshToken, saveAccountInfo } from './lib/token-store.js';
import { scopes, redirectUri, msalConfig } from './lib/config.js';
import type { PkceCodes } from './types.js';
import http from 'http';
import open from 'open';

const pca = new PublicClientApplication(msalConfig);
const cryptoProvider = new CryptoProvider();

// Store PKCE codes for the auth flow
let pkceCodes: PkceCodes | null = null;

async function authenticate(): Promise<void> {
  console.log('Starting browser-based authentication...');
  console.log('A browser window will open for you to sign in with your Microsoft account.\n');
  
  return new Promise((resolve, reject) => {
    // Create a local server to handle the OAuth redirect
    const httpServer = http.createServer(async (req, res) => {
      const url = new URL(req.url || '/', redirectUri);
      
      if (url.pathname === '/' && url.searchParams.has('code')) {
        const code = url.searchParams.get('code');
        
        try {
          // Exchange authorization code for token
          const tokenRequest = {
            code: code!,
            scopes: scopes,
            redirectUri: redirectUri,
            codeVerifier: pkceCodes!.verifier
          };
          
          const response: AuthenticationResult = await pca.acquireTokenByCode(tokenRequest);
          
          const expiresAt = response.expiresOn || new Date(Date.now() + 3600 * 1000);
          await saveToken(response.accessToken, expiresAt);
          
          // Save refresh token for persistent authentication
          if ((response as { refreshToken?: string }).refreshToken) {
            await saveRefreshToken((response as { refreshToken?: string }).refreshToken);
          }
          
          // Save account info for silent token acquisition
          if (response.account) {
            await saveAccountInfo(response.account);
          }
          
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
          
          httpServer.close();
          console.log('\nAuthentication successful!');
          console.log('Token stored securely in system credential manager.');
          console.log('Your authentication will persist across MCP sessions.');
          resolve();
          
        } catch (error) {
          const message = error instanceof Error ? error.message : 'Unknown error';
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
    <p style="color: #666; margin: 0 0 10px 0; font-size: 16px;">Error: ${message}</p>
    <p style="color: #999; margin: 0; font-size: 14px;">Please close this window and try again.</p>
  </div>
</body>
</html>`);
          httpServer.close();
          console.error('Failed to exchange code for token:', message);
          reject(error instanceof Error ? error : new Error(message));
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
        httpServer.close();
        console.error('Authentication error:', errorDescription);
        reject(new Error(errorDescription));
        
      } else if (url.pathname === '/favicon.ico') {
        res.writeHead(204);
        res.end();
      }
    });
    
    httpServer.listen(8400, async () => {
      console.log('Local server listening on http://localhost:8400');
      
      try {
        // Generate PKCE codes
        pkceCodes = await cryptoProvider.generatePkceCodes();
        
        const authCodeUrlParameters = {
          scopes: scopes,
          redirectUri: redirectUri,
          codeChallenge: pkceCodes.challenge,
          codeChallengeMethod: 'S256' as const
        };
        
        const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
        console.log('Opening browser for Microsoft login...');
        await open(authUrl);
        
      } catch (error) {
        httpServer.close();
        const message = error instanceof Error ? error.message : 'Unknown error';
        console.error('Failed to generate auth URL:', message);
        reject(error instanceof Error ? error : new Error(message));
      }
    });
    
    // Timeout after 2 minutes
    setTimeout(() => {
      httpServer.close();
      reject(new Error('Authentication timed out after 2 minutes'));
    }, 120000);
  });
}

authenticate().then(() => {
  process.exit(0);
}).catch((error: Error) => {
  console.error('Authentication failed:', error.message);
  process.exit(1);
});
