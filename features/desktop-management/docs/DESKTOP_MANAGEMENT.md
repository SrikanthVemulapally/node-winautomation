# Desktop Management API

This document describes the Child Session and RDP Desktop Management APIs for creating isolated Windows desktop environments within your application windows.

## Overview

The Desktop Management API allows you to:
- Create isolated Windows child sessions
- Embed remote desktops via RDP loopback into application windows
- Launch and manage processes in isolated sessions
- Control desktop size, color depth, and connection settings
- Monitor session state and running processes

## Architecture

```
Your Application Window (Electron/WinForms/WPF)
  └─> RDP Client Control (Embedded)
       └─> Localhost RDP Connection (127.0.0.1:3389)
            └─> Child Session (Isolated Desktop)
                 └─> Applications running in isolation
```

## Requirements

- **Windows 10/11 Pro or Enterprise** (Child sessions not available on Home editions)
- **RDP enabled** on the local machine
- **Administrator privileges** for creating child sessions
- **Terminal Services** must be running
- A parent window with a valid `HWND` handle for embedding

## API Reference

### DesktopManager

High-level API for managing desktop-in-window functionality.

#### Constructor

```javascript
const desktopManager = new automation.DesktopManager();
```

#### Methods

##### `createDesktopInWindow(windowHandle, width?, height?)`

Creates a child session and embeds it in the specified window via RDP.

**Parameters:**
- `windowHandle` (number | Buffer): Native window handle (HWND)
- `width` (number, optional): Desktop width in pixels (default: 1024)
- `height` (number, optional): Desktop height in pixels (default: 768)

**Returns:** `DesktopInWindowResult`
```typescript
{
    sessionId: number;
    childSession: ChildSession;
    rdpClient: RDPClient;
}
```

**Example:**
```javascript
const result = desktopManager.createDesktopInWindow(hwnd, 1280, 720);
console.log(`Session ID: ${result.sessionId}`);
```

##### `getChildSession()`

Returns the managed child session instance.

**Returns:** `ChildSession | null`

##### `getRDPClient()`

Returns the managed RDP client instance.

**Returns:** `RDPClient | null`

##### `launchApplication(processPath, commandLine?)`

Launches an application in the child session.

**Parameters:**
- `processPath` (string): Path to executable
- `commandLine` (string, optional): Command line arguments

**Returns:** `LaunchedProcessInfo`
```typescript
{
    processId: number;
    threadId: number;
}
```

**Example:**
```javascript
const proc = desktopManager.launchApplication('notepad.exe', 'C:\\file.txt');
console.log(`Started PID: ${proc.processId}`);
```

##### `resizeDesktop(width, height)`

Resizes the embedded desktop.

**Parameters:**
- `width` (number): New width in pixels
- `height` (number): New height in pixels

**Returns:** `boolean`

##### `cleanup()`

Disconnects RDP and terminates the child session.

**Returns:** `boolean`

##### `getSessionId()`

Returns the child session ID.

**Returns:** `number`

##### `isConnected()`

Checks if RDP connection is active.

**Returns:** `boolean`

---

### ChildSession

Low-level API for managing Windows child sessions.

#### Constructor

```javascript
// Create new child session
const childSession = new automation.ChildSession();

// Or attach to existing session
const childSession = new automation.ChildSession(sessionId);
```

#### Methods

##### `getSessionId()`

Returns the session ID.

**Returns:** `number`

##### `isActive()`

Checks if the session is still active.

**Returns:** `boolean`

##### `terminate()`

Terminates the child session and all its processes.

**Returns:** `boolean`

##### `getProcesses()`

Lists all processes running in the session.

**Returns:** `ProcessInfo[]`
```typescript
{
    processId: number;
    processName: string;
    sessionId: number;
}
```

**Example:**
```javascript
const processes = childSession.getProcesses();
processes.forEach(proc => {
    console.log(`${proc.processName} (PID: ${proc.processId})`);
});
```

##### `launchProcess(processPath, commandLine?)`

Launches a process in the child session.

**Parameters:**
- `processPath` (string): Path to executable
- `commandLine` (string, optional): Command line arguments

**Returns:** `LaunchedProcessInfo`

**Example:**
```javascript
const proc = childSession.launchProcess('calc.exe');
```

##### `getSessionInfo()`

Returns detailed session information.

**Returns:** `SessionInfo`
```typescript
{
    sessionId: number;
    userName?: string;
    domainName?: string;
    state: string; // "Active", "Connected", "Disconnected", etc.
}
```

---

### RDPClient

Low-level API for controlling RDP connections.

#### Constructor

```javascript
const rdpClient = new automation.RDPClient();
```

#### Methods

##### `connect()`

Initiates the RDP connection.

**Returns:** `boolean`

##### `disconnect()`

Disconnects the RDP session.

**Returns:** `boolean`

##### `setServer(server)`

Sets the RDP server address.

**Parameters:**
- `server` (string): Server address (use "localhost" for child sessions)

**Returns:** `boolean`

##### `setPort(port)`

Sets the RDP port.

**Parameters:**
- `port` (number): Port number (default: 3389)

**Returns:** `boolean`

##### `setDesktopSize(width, height)`

Sets the desktop dimensions.

**Parameters:**
- `width` (number): Width in pixels
- `height` (number): Height in pixels

**Returns:** `boolean`

##### `setFullscreen(fullscreen)`

Enables or disables fullscreen mode.

**Parameters:**
- `fullscreen` (boolean): Fullscreen state

**Returns:** `boolean`

##### `getConnectionState()`

Returns the current connection state.

**Returns:** `number`
- `0`: Disconnected
- `1`: Connecting
- `2`: Connected

##### `embedInWindow(windowHandle)`

Embeds the RDP control in a parent window.

**Parameters:**
- `windowHandle` (number | Buffer): Native window handle (HWND)

**Returns:** `boolean`

##### `sendKeys(keys)`

Sends keyboard input to the remote session.

**Parameters:**
- `keys` (string): Keys to send

**Returns:** `boolean`

##### `getWindowHandle()`

Returns the RDP control's window handle.

**Returns:** `number | null`

##### `setColorDepth(depth)`

Sets the color depth.

**Parameters:**
- `depth` (number): Color depth (8, 15, 16, 24, or 32)

**Returns:** `boolean`

##### `setAuthenticationLevel(level)`

Sets the authentication level.

**Parameters:**
- `level` (number): Authentication level (0 = no authentication required)

**Returns:** `boolean`

##### `enableCredentialSaving(enable)`

Enables or disables credential saving.

**Parameters:**
- `enable` (boolean): Enable credential saving

**Returns:** `boolean`

##### `setUsername(username)`

Sets the username for authentication.

**Parameters:**
- `username` (string): Username

**Returns:** `boolean`

##### `setDomain(domain)`

Sets the domain for authentication.

**Parameters:**
- `domain` (string): Domain name

**Returns:** `boolean`

##### `reconnect()`

Attempts to reconnect the RDP session.

**Returns:** `boolean`

---

## Usage Examples

### Basic Desktop-in-Window

```javascript
const automation = require('node-winautomation');

const desktopManager = new automation.DesktopManager();
const result = desktopManager.createDesktopInWindow(hwnd, 1024, 768);

// Launch application
desktopManager.launchApplication('notepad.exe');

// Cleanup when done
desktopManager.cleanup();
```

### Electron Integration

```javascript
const { BrowserWindow } = require('electron');
const automation = require('node-winautomation');

const win = new BrowserWindow({ width: 1280, height: 800 });
const hwnd = win.getNativeWindowHandle();

const desktopManager = new automation.DesktopManager();
const result = desktopManager.createDesktopInWindow(hwnd, 1200, 700);

// Handle window resize
win.on('resize', () => {
    const [width, height] = win.getSize();
    desktopManager.resizeDesktop(width - 80, height - 100);
});

// Cleanup on close
win.on('close', () => {
    desktopManager.cleanup();
});
```

### Manual Session Management

```javascript
const automation = require('node-winautomation');

// Create child session
const childSession = new automation.ChildSession();
console.log(`Session ID: ${childSession.getSessionId()}`);

// Launch applications
childSession.launchProcess('calc.exe');
childSession.launchProcess('notepad.exe');

// Monitor processes
const processes = childSession.getProcesses();
console.log(`Running processes: ${processes.length}`);

// Get session info
const info = childSession.getSessionInfo();
console.log(`State: ${info.state}`);

// Cleanup
childSession.terminate();
```

### Advanced RDP Configuration

```javascript
const automation = require('node-winautomation');

const rdpClient = new automation.RDPClient();

// Configure connection
rdpClient.setServer('localhost');
rdpClient.setPort(3389);
rdpClient.setDesktopSize(1920, 1080);
rdpClient.setColorDepth(32);
rdpClient.setFullscreen(false);
rdpClient.setAuthenticationLevel(0);

// Embed in window
rdpClient.embedInWindow(hwnd);

// Connect
rdpClient.connect();

// Send input
rdpClient.sendKeys('Hello World');

// Disconnect
rdpClient.disconnect();
```

## Best Practices

1. **Always cleanup**: Call `cleanup()` or `terminate()` when done to free resources
2. **Error handling**: Wrap calls in try-catch blocks
3. **Check privileges**: Ensure running with administrator rights
4. **Verify RDP**: Confirm RDP is enabled before creating sessions
5. **Window handles**: Ensure parent window is valid before embedding
6. **Session monitoring**: Regularly check `isActive()` to detect disconnections
7. **Resource limits**: Be mindful of system resources when creating multiple sessions

## Troubleshooting

### "Failed to create child session"
- Ensure running as Administrator
- Verify Windows edition (Pro/Enterprise required)
- Check that Terminal Services is running

### "Failed to connect" (RDP)
- Verify RDP is enabled in Windows settings
- Check firewall settings for port 3389
- Ensure localhost connections are allowed

### "Invalid window handle"
- Verify the window exists and is valid
- Use `IsWindow()` to check handle validity
- Ensure window is created before embedding

### Session terminates unexpectedly
- Check system event logs
- Verify sufficient system resources
- Monitor for policy restrictions

## Security Considerations

- Child sessions run with the same user privileges
- Applications in child sessions can access user data
- RDP connections use localhost but still follow RDP security policies
- Consider using authentication even for localhost connections
- Monitor and log session activities for audit purposes

## Performance Notes

- Each child session consumes system resources (RAM, CPU)
- RDP encoding/decoding has performance overhead
- Limit the number of concurrent sessions
- Use appropriate color depth (lower = better performance)
- Consider desktop size impact on performance

## See Also

- [COM API Documentation](./COM_API.md)
- [UI Automation Documentation](./GETTING_STARTED.md)
- [Examples](../examples/)
