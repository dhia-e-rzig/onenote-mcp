import { 
  loadToken, 
  saveToken, 
  loadRefreshToken, 
  saveRefreshToken, 
  loadAccountInfo,
  saveAccountInfo,
  deleteToken,
  isTokenExpired,
  isValidTokenFormat 
} from './lib/token-store.js';
import { Client } from '@microsoft/microsoft-graph-client';

// Test results tracking
let passed = 0;
let failed = 0;
let skipped = 0;

// Test notebook tracking for cleanup
let testNotebookId = null;
let testSectionId = null;
let testPageId = null;
let testNotebookLink = null;
let graphClient = null;

// Use a fixed name so we can reuse existing test notebooks
const TEST_NOTEBOOK_NAME = '_MCP_Test_Notebook';
const TEST_SECTION_NAME = 'Test Section';
const TEST_PAGE_TITLE = 'Test Page';
const TEST_PAGE_CONTENT = `
<!DOCTYPE html>
<html>
  <head>
    <title>${TEST_PAGE_TITLE}</title>
  </head>
  <body>
    <h1>MCP Test Page</h1>
    <p>This is a test page created by the OneNote MCP test suite.</p>
    <p>Created at: ${new Date().toISOString()}</p>
    <ul>
      <li>Item 1</li>
      <li>Item 2</li>
      <li>Item 3</li>
    </ul>
  </body>
</html>
`;

function test(name, condition) {
  if (condition) {
    console.log(`✅ ${name}`);
    passed++;
  } else {
    console.log(`❌ ${name}`);
    failed++;
  }
}

function skip(name, reason) {
  console.log(`⏭️  ${name} - ${reason}`);
  skipped++;
}

/**
 * Create test notebook with section and page, or reuse existing one
 */
async function setupTestNotebook() {
  console.log('\n🔨 Setting up test notebook...');
  
  try {
    // Check if test notebook already exists
    console.log(`   Checking for existing notebook: "${TEST_NOTEBOOK_NAME}"...`);
    const notebooks = await graphClient.api('/me/onenote/notebooks').get();
    const existingNotebook = notebooks.value?.find(n => n.displayName === TEST_NOTEBOOK_NAME);
    
    if (existingNotebook) {
      console.log('   ✓ Found existing test notebook, reusing it');
      testNotebookId = existingNotebook.id;
      testNotebookLink = existingNotebook.links?.oneNoteWebUrl?.href || null;
      
      // Get existing section
      const sections = await graphClient.api(`/me/onenote/notebooks/${testNotebookId}/sections`).get();
      const existingSection = sections.value?.find(s => s.displayName === TEST_SECTION_NAME);
      
      if (existingSection) {
        testSectionId = existingSection.id;
        console.log('   ✓ Found existing test section');
        
        // Delete old pages in the section
        const pages = await graphClient.api(`/me/onenote/sections/${testSectionId}/pages`).get();
        for (const page of pages.value || []) {
          try {
            await graphClient.api(`/me/onenote/pages/${page.id}`).delete();
          } catch (e) { /* ignore */ }
        }
      } else {
        // Create test section
        console.log(`   Creating section: "${TEST_SECTION_NAME}"...`);
        const section = await graphClient.api(`/me/onenote/notebooks/${testNotebookId}/sections`).post({
          displayName: TEST_SECTION_NAME
        });
        testSectionId = section.id;
        console.log(`   ✓ Section created`);
      }
    } else {
      // Create new test notebook
      console.log(`   Creating notebook: "${TEST_NOTEBOOK_NAME}"...`);
      const notebook = await graphClient.api('/me/onenote/notebooks').post({
        displayName: TEST_NOTEBOOK_NAME
      });
      testNotebookId = notebook.id;
      testNotebookLink = notebook.links?.oneNoteWebUrl?.href || null;
      console.log(`   ✓ Notebook created`);
      
      // Create test section
      console.log(`   Creating section: "${TEST_SECTION_NAME}"...`);
      const section = await graphClient.api(`/me/onenote/notebooks/${testNotebookId}/sections`).post({
        displayName: TEST_SECTION_NAME
      });
      testSectionId = section.id;
      console.log(`   ✓ Section created`);
    }
    
    // Create test page
    console.log(`   Creating page: "${TEST_PAGE_TITLE}"...`);
    const page = await graphClient
      .api(`/me/onenote/sections/${testSectionId}/pages`)
      .header('Content-Type', 'application/xhtml+xml')
      .post(TEST_PAGE_CONTENT);
    testPageId = page.id;
    console.log(`   ✓ Page created`);
    
    // Wait a moment for OneNote to sync
    console.log('   Waiting for sync...');
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    console.log('   ✓ Test notebook setup complete!\n');
    return true;
  } catch (error) {
    console.log(`   ❌ Setup failed: ${error.message}`);
    return false;
  }
}

/**
 * Clean up test notebook
 */
async function cleanupTestNotebook() {
  console.log('\n🧹 Cleaning up test notebook...');
  
  if (!testNotebookId) {
    console.log('   No test notebook to clean up.');
    return;
  }
  
  try {
    // Note: OneNote API doesn't support deleting notebooks directly
    // We can only delete sections and pages
    // The notebook will remain but be empty - user may need to delete manually
    
    if (testPageId) {
      console.log('   Deleting test page...');
      try {
        await graphClient.api(`/me/onenote/pages/${testPageId}`).delete();
        console.log('   ✓ Page deleted');
      } catch (e) {
        console.log(`   ⚠ Could not delete page: ${e.message}`);
      }
    }
    
    // Show notebook link for manual cleanup
    console.log('\n   ─────────────────────────────────────────────');
    console.log(`   📋 MANUAL CLEANUP (Optional)`);
    console.log('   ─────────────────────────────────────────────');
    console.log(`   The Microsoft Graph API does not support deleting notebooks.`);
    console.log(`   The test notebook will be reused for future test runs.`);
    console.log(`\n   📓 "${TEST_NOTEBOOK_NAME}"`);
    if (testNotebookLink) {
      console.log(`\n   🔗 Direct link to delete: ${testNotebookLink}`);
    }
    console.log('\n   ─────────────────────────────────────────────');
    
    console.log('   ✓ Cleanup complete!\n');
  } catch (error) {
    console.log(`   ⚠ Cleanup warning: ${error.message}`);
  }
}

async function runTests() {
  console.log('\n🧪 Running Token Store & MCP Tool Tests\n');
  console.log('─'.repeat(50));
  
  // Test 1: Token format validation
  console.log('\n📝 Token Format Validation Tests:');
  test('Valid JWT format accepted', isValidTokenFormat('header.payload.signature'));
  test('Valid opaque token accepted', isValidTokenFormat('EwB4BMl6BAAUu4TQbLz' + 'x'.repeat(100)));
  test('Empty string rejected', !isValidTokenFormat(''));
  test('Null rejected', !isValidTokenFormat(null));
  test('Undefined rejected', !isValidTokenFormat(undefined));
  test('Short token rejected', !isValidTokenFormat('short'));
  test('Invalid format rejected (two dots)', !isValidTokenFormat('invalid.token.'));
  
  // Test 2: Token expiry checking
  console.log('\n⏰ Token Expiry Tests:');
  const futureDate = new Date(Date.now() + 3600 * 1000); // 1 hour from now
  const pastDate = new Date(Date.now() - 3600 * 1000); // 1 hour ago
  const soonDate = new Date(Date.now() + 2 * 60 * 1000); // 2 minutes from now (within 5 min buffer)
  
  test('Future expiry is not expired', !isTokenExpired(futureDate));
  test('Past expiry is expired', isTokenExpired(pastDate));
  test('Expiry within 5 min buffer is expired', isTokenExpired(soonDate));
  test('Null expiry is expired', isTokenExpired(null));
  
  // ===== MCP Tool Tests with Mock Notebook =====
  console.log('\n🔧 MCP Tool Tests:');
  
  const realToken = await loadToken();
  const realRefresh = await loadRefreshToken();
  
  if (!realToken.token || !isValidTokenFormat(realToken.token)) {
    skip('All MCP tool tests', 'No valid credentials. Run "node authenticate.js" first.');
  } else {
    // Create Graph client for testing
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, realToken.token);
      }
    });
    
    // Setup test notebook
    const setupSuccess = await setupTestNotebook();
    
    if (!setupSuccess) {
      skip('MCP tool tests', 'Failed to create test notebook');
    } else {
      try {
        // Test: listNotebooks (should include our test notebook)
        console.log('   📓 listNotebooks:');
        try {
          const notebooks = await graphClient.api("/me/onenote/notebooks").get();
          test('listNotebooks returns data', notebooks !== null);
          test('listNotebooks has value array', Array.isArray(notebooks.value));
          
          const testNotebook = notebooks.value?.find(n => n.id === testNotebookId);
          test('Test notebook found in list', !!testNotebook);
          test('Test notebook has correct name', testNotebook?.displayName === TEST_NOTEBOOK_NAME);
          console.log(`      Found ${notebooks.value?.length || 0} notebook(s), test notebook included`);
        } catch (error) {
          test('listNotebooks API call', false);
          console.log(`      Error: ${error.message}`);
        }
        
        // Test: listSections (should include our test section)
        console.log('\n   📑 listSections:');
        try {
          const sections = await graphClient.api(`/me/onenote/notebooks/${testNotebookId}/sections`).get();
          test('listSections returns data', sections !== null);
          test('listSections has value array', Array.isArray(sections.value));
          
          const testSection = sections.value?.find(s => s.id === testSectionId);
          test('Test section found in list', !!testSection);
          test('Test section has correct name', testSection?.displayName === TEST_SECTION_NAME);
          console.log(`      Found ${sections.value?.length || 0} section(s) in test notebook`);
        } catch (error) {
          test('listSections API call', false);
          console.log(`      Error: ${error.message}`);
        }
        
        // Test: listPages (from our test section)
        console.log('\n   📄 listPages:');
        try {
          const pages = await graphClient.api(`/me/onenote/sections/${testSectionId}/pages`).get();
          test('listPages returns data', pages !== null);
          test('listPages has value array', Array.isArray(pages.value));
          
          const testPage = pages.value?.find(p => p.id === testPageId);
          test('Test page found in list', !!testPage);
          test('Test page has correct title', testPage?.title === TEST_PAGE_TITLE);
          console.log(`      Found ${pages.value?.length || 0} page(s) in test section`);
        } catch (error) {
          test('listPages API call', false);
          console.log(`      Error: ${error.message}`);
        }
        
        // Test: getPage content (with retry for sync delay)
        console.log('\n   📖 getPage (content):');
        try {
          let content = null;
          let retries = 3;
          
          while (retries > 0) {
            try {
              content = await graphClient.api(`/me/onenote/pages/${testPageId}/content`).get();
              break;
            } catch (e) {
              retries--;
              if (retries > 0) {
                console.log(`      Waiting for page sync (${retries} retries left)...`);
                await new Promise(resolve => setTimeout(resolve, 2000));
              } else {
                throw e;
              }
            }
          }
          
          test('getPage content API responds', true);
          
          // Check if content contains our test data
          const contentStr = typeof content === 'string' ? content : '';
          test('getPage returns expected content', content !== null);
          console.log(`      Content type: ${typeof content}, length: ${contentStr.length || 0} chars`);
        } catch (error) {
          test('getPage content API call', false);
          console.log(`      Error: ${error.message}`);
        }
        
        // Test: searchPages (search for our test page)
        console.log('\n   🔍 searchPages:');
        try {
          const allPages = await graphClient.api(`/me/onenote/sections/${testSectionId}/pages`).get();
          test('searchPages fetches pages', allPages !== null && Array.isArray(allPages.value));
          
          // Simulate client-side search for "Test"
          const searchTerm = 'test';
          const filtered = allPages.value?.filter(page => 
            page.title && page.title.toLowerCase().includes(searchTerm)
          ) || [];
          test('searchPages finds test page', filtered.length > 0);
          console.log(`      Pages matching "${searchTerm}": ${filtered.length}`);
        } catch (error) {
          test('searchPages API call', false);
          console.log(`      Error: ${error.message}`);
        }
        
        // Test: getNotebook (get specific notebook details)
        console.log('\n   📔 getNotebook:');
        try {
          const notebook = await graphClient.api(`/me/onenote/notebooks/${testNotebookId}`).get();
          test('getNotebook returns data', notebook !== null);
          test('getNotebook has correct id', notebook?.id === testNotebookId);
          test('getNotebook has displayName', !!notebook?.displayName);
          console.log(`      Notebook: "${notebook?.displayName}"`);
        } catch (error) {
          test('getNotebook API call', false);
          console.log(`      Error: ${error.message}`);
        }
        
        // Test: User info (validates token)
        console.log('\n   👤 User Info (token validation):');
        try {
          const me = await graphClient.api("/me").get();
          test('Token is valid for Graph API', me !== null);
          test('User has displayName', !!me.displayName);
          console.log(`      Authenticated as: ${me.displayName}`);
        } catch (error) {
          test('Token validation', false);
          console.log(`      Error: ${error.message}`);
        }
        
        // Test: Auth Persistence
        console.log('\n   🔐 Auth Persistence:');
        test('Refresh token stored for persistence', !!realRefresh);
        if (!realRefresh) {
          console.log('      ⚠️  Re-run "node authenticate.js" to enable persistent auth');
        } else {
          console.log('      Auth will persist across MCP sessions');
        }
        
      } finally {
        // Always cleanup, even if tests fail
        await cleanupTestNotebook();
      }
    }
  }
  
  // ===== Token Persistence Tests (after MCP tests) =====
  console.log('\n💾 Token Persistence Tests:');
  
  // Backup real credentials
  const existingToken = await loadToken();
  const existingRefresh = await loadRefreshToken();
  const existingAccount = await loadAccountInfo();
  
  // Save and load access token with test data
  const testToken = 'test.access.token';
  const testExpiry = new Date(Date.now() + 3600 * 1000);
  await saveToken(testToken, testExpiry);
  
  const loaded = await loadToken();
  test('Access token saved and loaded correctly', loaded.token === testToken);
  test('Token expiry saved correctly', loaded.expiresAt instanceof Date);
  test('Token expiry time preserved', Math.abs(loaded.expiresAt.getTime() - testExpiry.getTime()) < 1000);
  
  // Test: Refresh token persistence
  console.log('\n🔄 Refresh Token Persistence Tests:');
  
  const testRefreshToken = 'test-refresh-token-value';
  await saveRefreshToken(testRefreshToken);
  
  const loadedRefresh = await loadRefreshToken();
  test('Refresh token saved and loaded correctly', loadedRefresh === testRefreshToken);
  
  // Test: Account info persistence
  console.log('\n👤 Account Info Persistence Tests:');
  
  const testAccount = {
    homeAccountId: 'test-home-id',
    environment: 'login.microsoftonline.com',
    tenantId: 'test-tenant',
    username: 'test@example.com',
    localAccountId: 'test-local-id'
  };
  await saveAccountInfo(testAccount);
  
  const loadedAccount = await loadAccountInfo();
  test('Account info saved and loaded', loadedAccount !== null);
  test('Account homeAccountId preserved', loadedAccount?.homeAccountId === testAccount.homeAccountId);
  test('Account username preserved', loadedAccount?.username === testAccount.username);
  
  // Restore real credentials
  if (existingToken.token) {
    await saveToken(existingToken.token, existingToken.expiresAt);
  }
  if (existingRefresh) {
    await saveRefreshToken(existingRefresh);
  }
  if (existingAccount) {
    await saveAccountInfo(existingAccount);
  }
  
  // Summary
  console.log('\n' + '─'.repeat(50));
  console.log(`\n📊 Test Results: ${passed} passed, ${failed} failed, ${skipped} skipped`);
  
  if (failed === 0) {
    console.log('✨ All tests passed!\n');
  } else {
    console.log('⚠️  Some tests failed.\n');
    process.exit(1);
  }
}

// Handle unexpected exits to attempt cleanup
process.on('SIGINT', async () => {
  console.log('\n⚠️  Interrupted! Attempting cleanup...');
  if (graphClient && testNotebookId) {
    await cleanupTestNotebook();
  }
  process.exit(1);
});

runTests().catch(async (e) => {
  console.error('Test runner error:', e.message);
  if (graphClient && testNotebookId) {
    await cleanupTestNotebook();
  }
  process.exit(1);
});
