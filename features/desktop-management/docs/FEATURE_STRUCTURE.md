# Desktop Management Feature - Folder Structure

This document describes the organization of the Desktop Management feature as a separate, modular component.

## Rationale

The Desktop Management feature is organized as a standalone feature module to:
- **Maintain separation** from core UI Automation functionality
- **Enable optional inclusion** - can be excluded from builds if not needed
- **Simplify maintenance** - all related code in one location
- **Improve clarity** - clear distinction between features
- **Support modularity** - easy to add more features in the future

## Directory Structure

```
features/desktop-management/
│
├── DesktopManagement.h              # Feature header (includes all components)
├── README.md                        # Feature overview and quick start
├── IMPLEMENTATION_SUMMARY.md        # Technical implementation details
├── FEATURE_STRUCTURE.md            # This file
│
├── wrappers/                        # C++ wrapper implementations
│   ├── ChildSessionWrapper.h        # Child session management
│   ├── ChildSessionWrapper.cc
│   ├── RDPClientWrapper.h           # RDP client control
│   ├── RDPClientWrapper.cc
│   ├── DesktopManagerWrapper.h      # High-level unified API
│   └── DesktopManagerWrapper.cc
│
├── examples/                        # Feature-specific examples
│   └── desktop-in-window-example.js # Comprehensive usage examples
│
└── docs/                           # Feature documentation
    └── DESKTOP_MANAGEMENT.md        # Complete API reference
```

## Integration Points

### 1. Build System (`binding.gyp`)
```python
# Desktop Management Feature
"features/desktop-management/wrappers/ChildSessionWrapper.cc",
"features/desktop-management/wrappers/RDPClientWrapper.cc",
"features/desktop-management/wrappers/DesktopManagerWrapper.cc",
```

### 2. Library Header (`Library.h`)
```cpp
// Features
#include "features/desktop-management/DesktopManagement.h"
```

### 3. AutomationAddon (`AutomationAddon.h/cc`)
- Constructor references for feature wrappers
- Initialization in addon constructor
- JavaScript exports

### 4. TypeScript Definitions (`index.d.ts`)
- Interface definitions for all feature classes
- Type exports for JavaScript usage

## Dependencies

### System Libraries (Windows)
- `Wtsapi32.lib` - Windows Terminal Services API
- `Userenv.lib` - User environment management
- `ole32.lib` - COM/OLE base support
- `oleaut32.lib` - OLE Automation support

### Internal Dependencies
- `Shared.h` - Common types and includes
- `AutomationAddon` - Addon infrastructure

## Feature Components

### ChildSessionWrapper
**Purpose**: Manage Windows child sessions via WTS API

**Capabilities**:
- Create new child sessions
- Attach to existing sessions
- Launch processes in sessions
- Enumerate session processes
- Query session information
- Terminate sessions

### RDPClientWrapper
**Purpose**: Control RDP ActiveX client for remote desktop connections

**Capabilities**:
- Configure RDP connection parameters
- Embed RDP control in windows
- Connect/disconnect sessions
- Send keyboard input
- Query connection state
- Manage authentication

### DesktopManagerWrapper
**Purpose**: High-level API combining child session and RDP functionality

**Capabilities**:
- One-call desktop-in-window creation
- Automatic configuration
- Simplified application launching
- Desktop resizing
- Resource cleanup

## Usage Pattern

```javascript
// Import main library
const automation = require('node-winautomation');

// Desktop Management feature is automatically available
const desktopManager = new automation.DesktopManager();
const childSession = new automation.ChildSession();
const rdpClient = new automation.RDPClient();
```

## Separation from Core

### What's Separate
- ✅ All desktop management C++ code in `features/desktop-management/wrappers/`
- ✅ Feature-specific examples in `features/desktop-management/examples/`
- ✅ Feature documentation in `features/desktop-management/docs/`
- ✅ Feature overview in `features/desktop-management/README.md`

### What's Shared
- Build configuration (`binding.gyp`)
- AutomationAddon integration (`AutomationAddon.h/cc`)
- TypeScript definitions (`index.d.ts`)
- Main library header (`Library.h`)

## Adding New Features

To add a new feature module:

1. Create feature folder: `features/[feature-name]/`
2. Add subdirectories: `wrappers/`, `examples/`, `docs/`
3. Create feature header: `[FeatureName].h`
4. Update `binding.gyp` with source files
5. Include feature header in `Library.h`
6. Initialize in `AutomationAddon`
7. Add TypeScript definitions to `index.d.ts`
8. Document in feature README

## Build Configuration

The feature is included in the standard build process:

```bash
# Build with Desktop Management feature
npm run rebuild

# The feature is compiled as part of the main addon
# Output: build/Release/Automation.node
```

## Conditional Compilation (Future)

To make the feature truly optional, you could add:

```cpp
// In binding.gyp
"defines": [ "ENABLE_DESKTOP_MANAGEMENT" ]

// In code
#ifdef ENABLE_DESKTOP_MANAGEMENT
#include "features/desktop-management/DesktopManagement.h"
#endif
```

## Testing

Feature-specific tests should be placed in:
```
features/desktop-management/tests/
```

## Documentation Links

- **API Reference**: [docs/DESKTOP_MANAGEMENT.md](./docs/DESKTOP_MANAGEMENT.md)
- **Examples**: [examples/desktop-in-window-example.js](./examples/desktop-in-window-example.js)
- **Implementation**: [IMPLEMENTATION_SUMMARY.md](./IMPLEMENTATION_SUMMARY.md)
- **Main README**: [../../README.md](../../README.md)

## Maintenance

When updating the feature:
1. Keep all changes within `features/desktop-management/`
2. Update version in feature README if significant changes
3. Update API documentation if interfaces change
4. Add examples for new functionality
5. Update TypeScript definitions

## Benefits of This Structure

1. **Clear Boundaries**: Easy to see what belongs to the feature
2. **Independent Evolution**: Feature can be updated without affecting core
3. **Optional Inclusion**: Can be excluded from builds if needed
4. **Better Organization**: Related code grouped together
5. **Easier Testing**: Feature can be tested in isolation
6. **Documentation**: All feature docs in one place
7. **Examples**: Feature-specific examples don't clutter main examples
8. **Scalability**: Pattern can be repeated for other features

## Related Features (Future)

Potential future feature modules:
- `features/screen-capture/` - Screen recording and capture
- `features/input-simulation/` - Advanced input simulation
- `features/window-management/` - Window manipulation utilities
- `features/process-management/` - Process control and monitoring
