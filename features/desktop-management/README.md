# Desktop Management Feature

This feature provides Windows child session management and RDP desktop-in-window embedding capabilities.

## Overview

Create isolated Windows desktop environments and embed them inside your application windows using RDP loopback connections.

## Components

### Wrappers
- **ChildSessionWrapper** - Manages Windows child sessions via WTS API
- **RDPClientWrapper** - Controls RDP ActiveX client for remote desktop connections
- **DesktopManagerWrapper** - High-level unified API for desktop-in-window functionality

### Documentation
- [API Documentation](./docs/DESKTOP_MANAGEMENT.md) - Complete API reference
- [Implementation Summary](./IMPLEMENTATION_SUMMARY.md) - Technical details

### Examples
- [desktop-in-window-example.js](./examples/desktop-in-window-example.js) - Comprehensive usage examples

## Quick Start

```javascript
const automation = require('node-winautomation');

// Create desktop manager
const desktopManager = new automation.DesktopManager();

// Create desktop in window (e.g., Electron window handle)
const result = desktopManager.createDesktopInWindow(hwnd, 1280, 720);

// Launch applications in isolated desktop
desktopManager.launchApplication('notepad.exe');

// Cleanup when done
desktopManager.cleanup();
```

## Requirements

- Windows 10/11 Pro or Enterprise
- RDP enabled on local machine
- Administrator privileges
- Terminal Services running
- Valid parent window handle (HWND)

## Architecture

```
Application Window
    ↓
DesktopManager
    ↓
┌─────────────────┬──────────────────┐
│  ChildSession   │    RDPClient     │
└─────────────────┴──────────────────┘
    ↓                     ↓
Child Session ←─────RDP Loopback─────┘
```

## Build Configuration

This feature adds the following dependencies:
- `Wtsapi32.lib` - Windows Terminal Services
- `Userenv.lib` - User environment management
- `ole32.lib` - COM/OLE support
- `oleaut32.lib` - OLE Automation

## Files

```
features/desktop-management/
├── DesktopManagement.h          # Feature header
├── README.md                    # This file
├── IMPLEMENTATION_SUMMARY.md    # Technical details
├── wrappers/
│   ├── ChildSessionWrapper.h
│   ├── ChildSessionWrapper.cc
│   ├── RDPClientWrapper.h
│   ├── RDPClientWrapper.cc
│   ├── DesktopManagerWrapper.h
│   └── DesktopManagerWrapper.cc
├── examples/
│   └── desktop-in-window-example.js
└── docs/
    └── DESKTOP_MANAGEMENT.md
```

## Integration

The feature is integrated into the main library through:
1. Build configuration in `binding.gyp`
2. AutomationAddon initialization
3. TypeScript definitions in `index.d.ts`
4. JavaScript exports

## See Also

- [Main Documentation](../../docs/)
- [UI Automation Features](../../wrappers/)
- [COM Integration](../../src/com/)
