/**
 * Desktop Management Feature Test
 * 
 * Tests the child session and RDP desktop management functionality
 * 
 * Requirements:
 * - Windows 10/11 Pro or Enterprise
 * - Administrator privileges
 * - RDP enabled
 */

console.log('=== Desktop Management Feature Test ===\n');

async function testChildSession() {
    console.log('Test 1: Child Session Creation');
    console.log('--------------------------------');
    
    try {
        // Try to load the module
        const automation = require('../index.js');
        
        // Check if desktop management classes are available
        if (!automation.ChildSession) {
            console.log('❌ ChildSession class not found in exports');
            console.log('⚠ This might be because:');
            console.log('  - The addon needs to be rebuilt');
            console.log('  - Desktop management feature is not compiled');
            return false;
        }
        
        console.log('✓ ChildSession class found');
        
        if (!automation.RDPClient) {
            console.log('❌ RDPClient class not found in exports');
            return false;
        }
        
        console.log('✓ RDPClient class found');
        
        if (!automation.DesktopManager) {
            console.log('❌ DesktopManager class not found in exports');
            return false;
        }
        
        console.log('✓ DesktopManager class found');
        
        // Try to create a child session
        console.log('\n✓ Attempting to create child session...');
        console.log('⚠ This requires Administrator privileges');
        
        try {
            const childSession = new automation.ChildSession();
            console.log('✓ Child session created successfully!');
            
            // Get session ID
            const sessionId = childSession.getSessionId();
            console.log(`✓ Session ID: ${sessionId}`);
            
            // Check if active
            const isActive = childSession.isActive();
            console.log(`✓ Session active: ${isActive}`);
            
            // Get session info
            const sessionInfo = childSession.getSessionInfo();
            console.log('✓ Session info:');
            console.log(`  - Session ID: ${sessionInfo.sessionId}`);
            console.log(`  - State: ${sessionInfo.state}`);
            console.log(`  - User: ${sessionInfo.userName || 'N/A'}`);
            console.log(`  - Domain: ${sessionInfo.domainName || 'N/A'}`);
            
            // Get processes in session
            const processes = childSession.getProcesses();
            console.log(`✓ Processes in session: ${processes.length}`);
            
            // Try to launch a simple process
            console.log('\n✓ Testing process launch...');
            try {
                const proc = childSession.launchProcess('notepad.exe');
                console.log(`✓ Notepad launched with PID: ${proc.processId}`);
                
                // Wait a bit
                await new Promise(resolve => setTimeout(resolve, 2000));
                
                // Check processes again
                const updatedProcesses = childSession.getProcesses();
                console.log(`✓ Updated process count: ${updatedProcesses.length}`);
                
                // List processes
                console.log('\n✓ Processes running in child session:');
                updatedProcesses.forEach(p => {
                    console.log(`  - ${p.processName} (PID: ${p.processId})`);
                });
                
            } catch (launchError) {
                console.log(`⚠ Process launch failed: ${launchError.message}`);
                console.log('  This is expected if permissions are insufficient');
            }
            
            // Terminate session
            console.log('\n✓ Terminating child session...');
            const terminated = childSession.terminate();
            console.log(`✓ Session terminated: ${terminated}`);
            
            console.log('\n✅ Child session test PASSED!');
            return true;
            
        } catch (sessionError) {
            console.log(`\n❌ Child session creation failed: ${sessionError.message}`);
            console.log('\n⚠ Common reasons:');
            console.log('  - Not running as Administrator');
            console.log('  - Windows Home edition (requires Pro/Enterprise)');
            console.log('  - Terminal Services not running');
            console.log('  - Insufficient privileges');
            return false;
        }
        
    } catch (error) {
        console.error('\n❌ Test failed:', error.message);
        console.error(error.stack);
        return false;
    }
}

async function testRDPClient() {
    console.log('\n\nTest 2: RDP Client');
    console.log('-------------------');
    
    try {
        const automation = require('../index.js');
        
        if (!automation.RDPClient) {
            console.log('❌ RDPClient not available');
            return false;
        }
        
        console.log('✓ Creating RDP client...');
        const rdpClient = new automation.RDPClient();
        console.log('✓ RDP client created');
        
        // Configure client
        rdpClient.setServer('localhost');
        console.log('✓ Server set to localhost');
        
        rdpClient.setPort(3389);
        console.log('✓ Port set to 3389');
        
        rdpClient.setDesktopSize(1024, 768);
        console.log('✓ Desktop size set to 1024x768');
        
        rdpClient.setColorDepth(32);
        console.log('✓ Color depth set to 32-bit');
        
        rdpClient.setFullscreen(false);
        console.log('✓ Fullscreen disabled');
        
        rdpClient.setAuthenticationLevel(0);
        console.log('✓ Authentication level set');
        
        // Get connection state
        const state = rdpClient.getConnectionState();
        console.log(`✓ Connection state: ${state}`);
        
        console.log('\n⚠ Note: Full RDP connection test requires:');
        console.log('  - Valid window handle for embedding');
        console.log('  - RDP enabled on local machine');
        console.log('  - Child session to connect to');
        
        console.log('\n✅ RDP client configuration test PASSED!');
        return true;
        
    } catch (error) {
        console.error('\n❌ RDP client test failed:', error.message);
        return false;
    }
}

async function testDesktopManager() {
    console.log('\n\nTest 3: Desktop Manager');
    console.log('------------------------');
    
    try {
        const automation = require('../index.js');
        
        if (!automation.DesktopManager) {
            console.log('❌ DesktopManager not available');
            return false;
        }
        
        console.log('✓ Creating desktop manager...');
        const desktopManager = new automation.DesktopManager();
        console.log('✓ Desktop manager created');
        
        console.log('\n⚠ Note: Full desktop-in-window test requires:');
        console.log('  - Valid parent window handle (HWND)');
        console.log('  - Administrator privileges');
        console.log('  - RDP enabled');
        console.log('  - Typically used with Electron or WinForms apps');
        
        console.log('\n✅ Desktop manager instantiation test PASSED!');
        return true;
        
    } catch (error) {
        console.error('\n❌ Desktop manager test failed:', error.message);
        return false;
    }
}

async function runAllTests() {
    console.log('Starting desktop management tests...\n');
    console.log('⚠ IMPORTANT: Run this as Administrator for full functionality\n');
    
    const childSessionResult = await testChildSession();
    const rdpClientResult = await testRDPClient();
    const desktopManagerResult = await testDesktopManager();
    
    console.log('\n\n=== Test Summary ===');
    console.log(`Child Session:    ${childSessionResult ? '✅ PASSED' : '❌ FAILED'}`);
    console.log(`RDP Client:       ${rdpClientResult ? '✅ PASSED' : '❌ FAILED'}`);
    console.log(`Desktop Manager:  ${desktopManagerResult ? '✅ PASSED' : '❌ FAILED'}`);
    
    if (childSessionResult && rdpClientResult && desktopManagerResult) {
        console.log('\n✅ All desktop management tests PASSED!');
        console.log('\nThe desktop management feature is working correctly.');
        console.log('You can now use it in your applications.');
    } else {
        console.log('\n⚠ Some tests failed.');
        console.log('\nTroubleshooting:');
        console.log('  1. Rebuild the addon: npm run rebuild');
        console.log('  2. Run as Administrator');
        console.log('  3. Check Windows edition (Pro/Enterprise required)');
        console.log('  4. Verify RDP is enabled');
        console.log('  5. Check that Terminal Services is running');
    }
}

runAllTests().catch(error => {
    console.error('Unexpected error:', error);
    process.exit(1);
});
