import { InteractiveBrowserCredential } from '@azure/identity';
import { saveToken } from './lib/token-store.js';

// Client ID - use environment variable for production
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e';

// Warn if using default client ID
if (!process.env.AZURE_CLIENT_ID) {
  console.log('WARNING: Using default Microsoft Graph Explorer client ID.');
  console.log('For production, set AZURE_CLIENT_ID environment variable.\n');
}

const scopes = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read'];

async function authenticate() {
  try {
    const credential = new InteractiveBrowserCredential({
      clientId: clientId,
      redirectUri: 'http://localhost:8400'
    });

    console.log('Starting authentication...');
    console.log('A browser window will open for you to sign in with your Microsoft account...\n');
    
    const tokenResponse = await credential.getToken(scopes);
    
    const expiresAt = tokenResponse.expiresOnTimestamp 
      ? new Date(tokenResponse.expiresOnTimestamp)
      : new Date(Date.now() + 3600 * 1000);
    
    await saveToken(tokenResponse.token, expiresAt);
    
    console.log('\nAuthentication successful!');
    console.log('Token stored securely in system credential manager.');
    
  } catch (error) {
    console.error('Authentication error:', error.message);
    process.exit(1);
  }
}

authenticate(); 