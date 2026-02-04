import { 
  loadToken, 
  saveToken, 
  loadRefreshToken, 
  saveRefreshToken, 
  loadAccountInfo,
  saveAccountInfo,
  isTokenExpired,
  isValidTokenFormat 
} from '../lib/token-store.js';

describe('Token Format Validation', () => {
  test('accepts valid JWT format (3 parts)', () => {
    expect(isValidTokenFormat('header.payload.signature')).toBe(true);
  });

  test('accepts valid opaque token (long string)', () => {
    expect(isValidTokenFormat('EwB4BMl6BAAUu4TQbLz' + 'x'.repeat(100))).toBe(true);
  });

  test('rejects empty string', () => {
    expect(isValidTokenFormat('')).toBe(false);
  });

  test('rejects null', () => {
    expect(isValidTokenFormat(null)).toBe(false);
  });

  test('rejects undefined', () => {
    expect(isValidTokenFormat(undefined)).toBe(false);
  });

  test('rejects short token', () => {
    expect(isValidTokenFormat('short')).toBe(false);
  });

  test('rejects invalid format with empty part', () => {
    expect(isValidTokenFormat('invalid.token.')).toBe(false);
  });
});

describe('Token Expiry Checking', () => {
  test('future expiry is not expired', () => {
    const futureDate = new Date(Date.now() + 3600 * 1000); // 1 hour from now
    expect(isTokenExpired(futureDate)).toBe(false);
  });

  test('past expiry is expired', () => {
    const pastDate = new Date(Date.now() - 3600 * 1000); // 1 hour ago
    expect(isTokenExpired(pastDate)).toBe(true);
  });

  test('expiry within 5 minute buffer is considered expired', () => {
    const soonDate = new Date(Date.now() + 2 * 60 * 1000); // 2 minutes from now
    expect(isTokenExpired(soonDate)).toBe(true);
  });

  test('null expiry is considered expired', () => {
    expect(isTokenExpired(null)).toBe(true);
  });
});

describe('Token Persistence', () => {
  let originalToken: { token: string | null; expiresAt: Date | null };
  let originalRefresh: string | null;
  let originalAccount: object | null;

  beforeAll(async () => {
    // Backup existing credentials
    originalToken = await loadToken();
    originalRefresh = await loadRefreshToken();
    originalAccount = await loadAccountInfo();
  });

  afterAll(async () => {
    // Restore original credentials
    if (originalToken.token) {
      await saveToken(originalToken.token, originalToken.expiresAt);
    }
    if (originalRefresh) {
      await saveRefreshToken(originalRefresh);
    }
    if (originalAccount) {
      await saveAccountInfo(originalAccount);
    }
  });

  test('saves and loads access token correctly', async () => {
    const testToken = 'test.access.token';
    const testExpiry = new Date(Date.now() + 3600 * 1000);
    
    await saveToken(testToken, testExpiry);
    const loaded = await loadToken();
    
    expect(loaded.token).toBe(testToken);
  });

  test('saves and loads token expiry correctly', async () => {
    const testToken = 'test.access.token';
    const testExpiry = new Date(Date.now() + 3600 * 1000);
    
    await saveToken(testToken, testExpiry);
    const loaded = await loadToken();
    
    expect(loaded.expiresAt).toBeInstanceOf(Date);
    expect(Math.abs(loaded.expiresAt!.getTime() - testExpiry.getTime())).toBeLessThan(1000);
  });

  test('saves and loads refresh token correctly', async () => {
    const testRefreshToken = 'test-refresh-token-value';
    
    await saveRefreshToken(testRefreshToken);
    const loadedRefresh = await loadRefreshToken();
    
    expect(loadedRefresh).toBe(testRefreshToken);
  });

  test('saves and loads account info correctly', async () => {
    const testAccount = {
      homeAccountId: 'test-home-id',
      environment: 'login.microsoftonline.com',
      tenantId: 'test-tenant',
      username: 'test@example.com',
      localAccountId: 'test-local-id'
    };
    
    await saveAccountInfo(testAccount);
    const loadedAccount = await loadAccountInfo();
    
    expect(loadedAccount).not.toBeNull();
    expect(loadedAccount?.homeAccountId).toBe(testAccount.homeAccountId);
    expect(loadedAccount?.username).toBe(testAccount.username);
  });
});
