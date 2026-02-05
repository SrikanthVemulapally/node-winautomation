# node-winautomation

A Node.js native addon for Windows UI Automation, providing comprehensive bindings to the Windows UI Automation API for automating Windows desktop applications.

## Overview

This project is based on [@bright-fish/node-ui-automation](https://github.com/bright-fish/node-ui-automation) and provides full access to Windows UI Automation APIs through native C++ bindings. It enables you to:

- Automate Windows desktop applications
- Inspect UI elements and their properties
- Interact with controls (buttons, text boxes, menus, etc.)
- Handle UI events and automation patterns
- Access accessibility information

## Features

- **Complete UI Automation Support**: Full bindings to Windows UI Automation API
- **Pattern Support**: All major automation patterns (Invoke, Value, Text, Selection, Grid, etc.)
- **Event Handling**: Support for automation events, focus changes, property changes, and structure changes
- **Type Definitions**: Full TypeScript definitions included
- **Native Performance**: C++ native addon for optimal performance

## Prerequisites

- **Windows Operating System** (Required)
- **Node.js** (v12 or higher recommended)
- **Visual Studio** or **Visual Studio Build Tools** with C++ support
- **Python** (for node-gyp)
- **node-gyp** installed globally: `npm install -g node-gyp`

## Installation

1. Clone or download this repository
2. Navigate to the project directory
3. Install dependencies and build:

```bash
npm install
npm run configure
npm run build
```

## Build Scripts

- `npm run configure` - Configure the native addon build
- `npm run build` - Build the native addon
- `npm run clean` - Clean build artifacts
- `npm run rebuild` - Clean and rebuild
- `npm test` - Run tests (from tests directory)

## Usage

### Basic Example - Click OK Button in About Windows

```javascript
const { Automation, PropertyIds, TreeScopes, PatternIds } = require('node-winautomation');

const automation = new Automation();

// Get desktop root element
const desktopElement = automation.getRootElement();

// Find "About Windows" window
const aboutWindowsProperty = automation.createPropertyCondition(
  PropertyIds.NamePropertyId,
  'About Windows'
);

const winverWindow = desktopElement.findFirst(
  TreeScopes.Subtree,
  aboutWindowsProperty
);

// Find OK button
const okProperty = automation.createPropertyCondition(
  PropertyIds.NamePropertyId,
  "OK"
);

const okButton = winverWindow.findFirst(
  TreeScopes.Subtree,
  okProperty
);

// Click the button
const invokeProvider = okButton.getCurrentPattern(
  PatternIds.InvokePatternId
);
invokeProvider.invoke();
```

### Advanced Example - Automate Notepad

```javascript
const { Automation, PropertyIds, ControlTypeIds, TreeScopes, PatternIds } = require('node-winautomation');

const automation = new Automation();
const root = automation.getRootElement();

// Find Notepad window
const condition = automation.createPropertyCondition(
  PropertyIds.NamePropertyId,
  'Notepad'
);

const notepadWindow = root.findFirst(
  TreeScopes.Children,
  condition
);

if (notepadWindow) {
  console.log('Found Notepad window');

  // Find the edit control
  const editCondition = automation.createPropertyCondition(
    PropertyIds.ControlTypePropertyId,
    ControlTypeIds.Edit
  );

  const editControl = notepadWindow.findFirst(
    TreeScopes.Descendants,
    editCondition
  );

  if (editControl) {
    // Use Value pattern to set text
    const valuePattern = editControl.getCurrentPattern(
      PatternIds.ValuePatternId
    );
    valuePattern.setValue('Hello from node-winautomation!');
  }
}
```

## API Reference

### Main Automation Object

The module exports the main `Automation` class which provides:

- `getRootElement()` - Get the desktop root element
- `createPropertyCondition(propertyId, value)` - Create search conditions
- `createTrueCondition()` / `createFalseCondition()` - Create boolean conditions
- `createAndCondition()` / `createOrCondition()` / `createNotCondition()` - Combine conditions
- Element walking and tree navigation methods

### Element Properties

Elements expose properties such as:
- `currentName` - Element name
- `currentClassName` - Class name
- `currentControlType` - Control type ID
- `currentBoundingRectangle` - Screen coordinates
- `currentIsEnabled` - Enabled state
- And many more...

### Automation Patterns

Supported patterns include:
- **Invoke Pattern** - Click buttons and menu items
- **Value Pattern** - Get/set values in text boxes
- **Text Pattern** - Advanced text manipulation
- **Selection Pattern** - Handle list/combo box selections
- **Toggle Pattern** - Work with checkboxes
- **Window Pattern** - Window manipulation
- **Grid Pattern** - Table/grid interaction
- And many more...

### Enumerations

The module provides access to all UI Automation enumerations:
- `ControlTypeIds` - Button, Edit, Window, etc.
- `PropertyIds` - Name, ClassName, ControlType, etc.
- `PatternIds` - Invoke, Value, Text, etc.
- `TreeScopes` - Element, Children, Descendants, Subtree, etc.

## Project Structure

```
node-winautomation/
├── AutomationAddon.cc/h      # Main addon entry point
├── binding.gyp                # Build configuration
├── index.js                   # Module entry
├── index.d.ts                 # TypeScript definitions
├── enumerations/              # UI Automation enumerations
├── patterns/                  # Automation pattern wrappers
├── utilities/                 # Helper utilities and event handlers
├── wrappers/                  # COM object wrappers
└── tests/                     # Test suite
```

## Testing

Tests are located in the `tests` directory:

```bash
cd tests
npm install
npm test
```

## Troubleshooting

### Build Errors

1. **"Cannot find module 'node-gyp'"**
   - Install node-gyp globally: `npm install -g node-gyp`

2. **"MSBuild.exe not found"**
   - Install Visual Studio Build Tools with C++ support
   - Or use: `npm install --global windows-build-tools`

3. **Python errors**
   - Ensure Python is installed and in PATH
   - node-gyp requires Python 2.7 or 3.x

### Runtime Errors

1. **"The specified module could not be found"**
   - Ensure the addon was built successfully
   - Check that `build/Release/Automation.node` exists

2. **Access Denied errors**
   - Some applications may require elevated privileges
   - Run Node.js as Administrator if needed

## Credits

This project is based on the excellent work by Chris Nimmons and Jake Cyr in the [@bright-fish/node-ui-automation](https://github.com/bright-fish/node-ui-automation) project.

## License

MIT License - See LICENSE file for details

## Resources

- [Windows UI Automation Overview](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)
- [UI Automation Control Patterns](https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpatternsoverview)
- [Node.js Addons Documentation](https://nodejs.org/api/addons.html)
