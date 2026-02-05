const automation = require('..');

/**
 * Desktop-in-Window Example
 * 
 * This example demonstrates how to create a child session and display it
 * inside a window using RDP loopback connection.
 * 
 * Requirements:
 * - Windows 10/11 Pro or Enterprise
 * - RDP must be enabled
 * - Administrator privileges
 * - A parent window to embed the desktop (e.g., Electron, WinForms, etc.)
 */

// Example 1: Using DesktopManager (Simplified API)
function createDesktopWithManager() {
    console.log('=== Desktop Manager Example ===\n');
    
    // Create a desktop manager instance
    const desktopManager = new automation.DesktopManager();
    
    // For this example, we'll use a dummy window handle
    // In a real application, you would get this from Electron's getNativeWindowHandle()
    // or from a WinForms/WPF window
    const dummyWindowHandle = 0; // Replace with actual window handle
    
    try {
        // Create a desktop in the specified window
        // This will:
        // 1. Create a child session
        // 2. Set up RDP client
        // 3. Connect to localhost
        // 4. Embed the RDP view in the window
        const result = desktopManager.createDesktopInWindow(
            dummyWindowHandle,
            1024,  // width
            768    // height
        );
        
        console.log(`Child session created with ID: ${result.sessionId}`);
        console.log('Desktop is now embedded in the window\n');
        
        // Launch an application in the child session
        console.log('Launching Notepad in child session...');
        const processInfo = desktopManager.launchApplication('notepad.exe');
        console.log(`Notepad started with PID: ${processInfo.processId}\n`);
        
        // Wait a bit
        setTimeout(() => {
            // Get session info
            const session = desktopManager.getChildSession();
            if (session) {
                const sessionInfo = session.getSessionInfo();
                console.log('Session Info:', sessionInfo);
                
                // List all processes in the session
                const processes = session.getProcesses();
                console.log(`\nProcesses in child session (${processes.length}):`);
                processes.forEach(proc => {
                    console.log(`  - ${proc.processName} (PID: ${proc.processId})`);
                });
            }
            
            // Resize the desktop
            console.log('\nResizing desktop to 1280x720...');
            desktopManager.resizeDesktop(1280, 720);
            
            // Check connection status
            console.log(`Connected: ${desktopManager.isConnected()}`);
            
            // Cleanup after 30 seconds
            setTimeout(() => {
                console.log('\nCleaning up...');
                desktopManager.cleanup();
                console.log('Desktop session terminated');
            }, 30000);
            
        }, 5000);
        
    } catch (error) {
        console.error('Error creating desktop:', error.message);
    }
}

// Example 2: Manual Control (Advanced API)
function createDesktopManually() {
    console.log('=== Manual Desktop Creation Example ===\n');
    
    try {
        // Step 1: Create a child session
        console.log('Creating child session...');
        const childSession = new automation.ChildSession();
        const sessionId = childSession.getSessionId();
        console.log(`Child session created with ID: ${sessionId}\n`);
        
        // Step 2: Create RDP client
        console.log('Creating RDP client...');
        const rdpClient = new automation.RDPClient();
        
        // Step 3: Configure RDP client
        rdpClient.setServer('localhost');
        rdpClient.setPort(3389);
        rdpClient.setDesktopSize(1024, 768);
        rdpClient.setColorDepth(32); // 32-bit color
        rdpClient.setFullscreen(false);
        rdpClient.setAuthenticationLevel(0); // Allow connections without authentication
        
        console.log('RDP client configured\n');
        
        // Step 4: Embed in window (use actual window handle)
        const dummyWindowHandle = 0; // Replace with actual window handle
        // rdpClient.embedInWindow(dummyWindowHandle);
        
        // Step 5: Connect
        console.log('Connecting to child session via RDP...');
        rdpClient.connect();
        console.log('Connected!\n');
        
        // Step 6: Launch applications
        console.log('Launching applications in child session...');
        
        const calc = childSession.launchProcess('calc.exe');
        console.log(`Calculator started (PID: ${calc.processId})`);
        
        const notepad = childSession.launchProcess('notepad.exe');
        console.log(`Notepad started (PID: ${notepad.processId})\n`);
        
        // Step 7: Monitor session
        setTimeout(() => {
            const sessionInfo = childSession.getSessionInfo();
            console.log('Session Information:');
            console.log(`  Session ID: ${sessionInfo.sessionId}`);
            console.log(`  User: ${sessionInfo.userName || 'N/A'}`);
            console.log(`  Domain: ${sessionInfo.domainName || 'N/A'}`);
            console.log(`  State: ${sessionInfo.state}\n`);
            
            const processes = childSession.getProcesses();
            console.log(`Active processes: ${processes.length}`);
            
            // Check if session is still active
            console.log(`Session active: ${childSession.isActive()}`);
            console.log(`RDP connection state: ${rdpClient.getConnectionState()}\n`);
            
            // Cleanup
            setTimeout(() => {
                console.log('Disconnecting and terminating session...');
                rdpClient.disconnect();
                childSession.terminate();
                console.log('Done!');
            }, 25000);
            
        }, 5000);
        
    } catch (error) {
        console.error('Error:', error.message);
    }
}

// Example 3: Electron Integration Example
function electronIntegrationExample() {
    console.log('=== Electron Integration Example ===\n');
    console.log('In your Electron app:\n');
    
    const exampleCode = `
// In Electron main process
const { BrowserWindow } = require('electron');
const automation = require('node-winautomation');

// Create Electron window
const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
        nodeIntegration: true,
        contextIsolation: false
    }
});

// Get native window handle
const hwnd = win.getNativeWindowHandle();

// Create desktop manager
const desktopManager = new automation.DesktopManager();

// Create desktop in Electron window
const result = desktopManager.createDesktopInWindow(
    hwnd,
    1200,  // width (slightly smaller than window)
    700    // height (slightly smaller than window)
);

console.log('Child desktop embedded in Electron window!');
console.log('Session ID:', result.sessionId);

// Launch application in child desktop
desktopManager.launchApplication('explorer.exe');

// Handle window resize
win.on('resize', () => {
    const [width, height] = win.getSize();
    desktopManager.resizeDesktop(width - 80, height - 100);
});

// Cleanup on close
win.on('close', () => {
    desktopManager.cleanup();
});
`;
    
    console.log(exampleCode);
}

// Example 4: Session Management
function sessionManagementExample() {
    console.log('=== Session Management Example ===\n');
    
    try {
        const childSession = new automation.ChildSession();
        const sessionId = childSession.getSessionId();
        
        console.log(`Created session: ${sessionId}\n`);
        
        // Launch multiple applications
        const apps = [
            { path: 'notepad.exe', name: 'Notepad' },
            { path: 'calc.exe', name: 'Calculator' },
            { path: 'mspaint.exe', name: 'Paint' }
        ];
        
        console.log('Launching applications...');
        apps.forEach(app => {
            try {
                const proc = childSession.launchProcess(app.path);
                console.log(`  ✓ ${app.name} (PID: ${proc.processId})`);
            } catch (err) {
                console.log(`  ✗ ${app.name} failed: ${err.message}`);
            }
        });
        
        // Monitor processes
        console.log('\nMonitoring processes...');
        const interval = setInterval(() => {
            if (!childSession.isActive()) {
                console.log('Session is no longer active');
                clearInterval(interval);
                return;
            }
            
            const processes = childSession.getProcesses();
            console.log(`Active processes: ${processes.length}`);
            
        }, 3000);
        
        // Terminate after 20 seconds
        setTimeout(() => {
            clearInterval(interval);
            console.log('\nTerminating session...');
            childSession.terminate();
            console.log('Session terminated');
        }, 20000);
        
    } catch (error) {
        console.error('Error:', error.message);
    }
}

// Run examples
console.log('Node.js Windows Automation - Desktop-in-Window Examples\n');
console.log('========================================================\n');

// Uncomment the example you want to run:

// createDesktopWithManager();
// createDesktopManually();
// electronIntegrationExample();
// sessionManagementExample();

console.log('\nNote: These examples require:');
console.log('  - Windows 10/11 Pro or Enterprise');
console.log('  - RDP enabled');
console.log('  - Administrator privileges');
console.log('  - A valid window handle for embedding\n');
