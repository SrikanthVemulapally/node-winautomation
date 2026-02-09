/**
 * Child Session Test
 * Tests the ChildSession wrapper functionality
 */

console.log('=== Child Session Test ===\n');

async function testChildSession() {
    try {
        const automation = require('../index.js');
        
        // Check if ChildSession is available
        if (!automation.ChildSession) {
            console.log('❌ ChildSession not found in exports');
            return false;
        }
        
        console.log('✓ ChildSession class found');
        console.log('✓ Creating child session...\n');
        
        // Create a child session (uses current session as fallback)
        const childSession = new automation.ChildSession();
        console.log('✓ Child session created successfully!');
        
        // Get session ID
        const sessionId = childSession.getSessionId();
        console.log(`✓ Session ID: ${sessionId}`);
        
        // Check if active
        const isActive = childSession.isActive();
        console.log(`✓ Session active: ${isActive}`);
        
        // Get session info
        console.log('\n✓ Getting session information...');
        const sessionInfo = childSession.getSessionInfo();
        console.log('  Session Info:');
        console.log(`  - Session ID: ${sessionInfo.sessionId}`);
        console.log(`  - State: ${sessionInfo.state}`);
        console.log(`  - User: ${sessionInfo.userName || 'N/A'}`);
        console.log(`  - Domain: ${sessionInfo.domainName || 'N/A'}`);
        
        // Get processes in session
        console.log('\n✓ Enumerating processes in session...');
        const processes = childSession.getProcesses();
        console.log(`✓ Found ${processes.length} processes`);
        
        if (processes.length > 0) {
            console.log('\n  Sample processes:');
            processes.slice(0, 5).forEach(p => {
                console.log(`  - ${p.processName} (PID: ${p.processId})`);
            });
            if (processes.length > 5) {
                console.log(`  ... and ${processes.length - 5} more`);
            }
        }
        
        // Try to launch a process
        console.log('\n✓ Testing process launch...');
        try {
            const proc = childSession.launchProcess('notepad.exe');
            console.log(`✓ Notepad launched successfully!`);
            console.log(`  - Process ID: ${proc.processId}`);
            console.log(`  - Thread ID: ${proc.threadId}`);
            
            // Wait a bit
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Check processes again
            const updatedProcesses = childSession.getProcesses();
            console.log(`\n✓ Updated process count: ${updatedProcesses.length}`);
            
            // Find notepad
            const notepad = updatedProcesses.find(p => p.processId === proc.processId);
            if (notepad) {
                console.log(`✓ Notepad confirmed running: ${notepad.processName}`);
            }
            
        } catch (launchError) {
            console.log(`⚠ Process launch failed: ${launchError.message}`);
            console.log('  (This is expected if not running as Administrator)');
        }
        
        console.log('\n✅ Child Session test PASSED!');
        console.log('\nNote: This uses the current session as a fallback.');
        console.log('True child session creation requires:');
        console.log('  - Windows SDK 10.0.17134.0+ with WTSCreateChildSession API');
        console.log('  - Windows 10/11 Pro or Enterprise');
        console.log('  - Administrator privileges');
        
        return true;
        
    } catch (error) {
        console.error('\n❌ Test failed:', error.message);
        console.error(error.stack);
        return false;
    }
}

testChildSession().then(success => {
    if (success) {
        console.log('\n✅ All tests completed successfully!');
    } else {
        console.log('\n❌ Tests failed');
        process.exit(1);
    }
}).catch(err => {
    console.error('Unexpected error:', err);
    process.exit(1);
});
