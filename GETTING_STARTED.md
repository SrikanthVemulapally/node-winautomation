# Getting Started with node-winautomation

This guide will help you get started with node-winautomation step by step.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation](#installation)
3. [Building the Project](#building-the-project)
4. [Running Examples](#running-examples)
5. [Writing Your First Automation](#writing-your-first-automation)
6. [Common Issues](#common-issues)

## Prerequisites

Before you begin, ensure you have the following installed:

### 1. Node.js

Download and install Node.js from [nodejs.org](https://nodejs.org/)

Verify installation:
```bash
node --version
npm --version
```

### 2. Visual Studio Build Tools

You need C++ build tools to compile the native addon.

**Option A: Full Visual Studio (Recommended)**
- Download [Visual Studio Community](https://visualstudio.microsoft.com/downloads/)
- During installation, select "Desktop development with C++"

**Option B: Build Tools Only**
- Download [Build Tools for Visual Studio](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022)
- Select "C++ build tools" workload

### 3. Python

node-gyp requires Python 2.7 or 3.x

Download from [python.org](https://www.python.org/downloads/)

Verify installation:
```bash
python --version
```

### 4. node-gyp

Install node-gyp globally:
```bash
npm install -g node-gyp
```

## Installation

1. **Navigate to the project directory**:
   ```bash
   cd C:\Users\Srika\CascadeProjects\node-winautomation
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

   This will install:
   - `bindings` - Helper for loading native addons
   - `node-addon-api` - N-API wrapper for building addons

## Building the Project

Build the native addon using these steps:

1. **Configure the build**:
   ```bash
   npm run configure
   ```

2. **Build the addon**:
   ```bash
   npm run build
   ```

   This creates `build/Release/Automation.node`

3. **Verify the build**:
   ```bash
   node -e "console.log(require('./index'))"
   ```

   You should see an object with UI Automation bindings.

### Alternative: Single Command

You can also rebuild in one step:
```bash
npm run rebuild
```

## Running Examples

The `examples/` directory contains sample automation scripts.

### Example 1: Notepad Automation

1. **Open Notepad** (notepad.exe)

2. **Run the example**:
   ```bash
   node examples/notepad-example.js
   ```

   This will automatically type text into Notepad.

### Example 2: Calculator Automation

1. **Open Calculator** (Search "Calculator" in Start menu)

2. **Run the example**:
   ```bash
   node examples/calculator-example.js
   ```

   This will perform a calculation (2 + 3 = 5).

## Writing Your First Automation

Create a new file `my-first-automation.js`:

```javascript
const { Automation, PropertyIds, TreeScopes } = require('node-winautomation');

// Initialize automation
const automation = new Automation();

// Get the desktop (root element)
const desktop = automation.getRootElement();

// Find all top-level windows
console.log('Top-level windows:');

const trueCondition = automation.createTrueCondition();
const allWindows = desktop.findAll(
  TreeScopes.Children,
  trueCondition
);

const windowCount = allWindows.length;
console.log(`Found ${windowCount} windows\n`);

for (let i = 0; i < windowCount; i++) {
  const window = allWindows.getElement(i);
  const name = window.currentName;
  const className = window.currentClassName;

  console.log(`${i + 1}. ${name || '(No name)'}`);
  console.log(`   Class: ${className || '(No class)'}\n`);
}

console.log('Done!');
```

Run it:
```bash
node my-first-automation.js
```

This will list all open windows on your desktop!

## Understanding the Basics

### The Automation Object

```javascript
const automation = new Automation();
```

This is your main entry point to the UI Automation API.

### The Root Element (Desktop)

```javascript
const desktop = automation.getRootElement();
```

The root element represents the entire desktop. All windows are children of this element.

### Finding Elements

You can find elements using conditions:

```javascript
// Find by name
const condition = automation.createPropertyCondition(
  PropertyIds.NamePropertyId,
  'Window Title'
);

const element = desktop.findFirst(
  TreeScopes.Children,    // or TreeScopes.Descendants for deep search
  condition
);
```

### Tree Scopes

- `TreeScopes.Element` - Only the element itself
- `TreeScopes.Children` - Direct children only
- `TreeScopes.Descendants` - All descendants (recursive)
- `TreeScopes.Subtree` - Element + all descendants

### Element Properties

Once you have an element, you can access its properties:

```javascript
console.log(element.currentName);
console.log(element.currentClassName);
console.log(element.currentControlType);
console.log(element.currentBoundingRectangle);
```

### Automation Patterns

Patterns let you interact with elements:

```javascript
// Invoke pattern (for buttons)
const invokePattern = button.getCurrentPattern(PatternIds.InvokePatternId);
invokePattern.invoke();

// Value pattern (for text boxes)
const valuePattern = textBox.getCurrentPattern(PatternIds.ValuePatternId);
valuePattern.setValue('Hello World');
const text = valuePattern.currentValue;
```

## Common Issues

### Build Errors

**Error: "Cannot find module 'node-gyp'"**
- Solution: `npm install -g node-gyp`

**Error: "MSBuild.exe not found"**
- Solution: Install Visual Studio Build Tools with C++ support

**Error: "Python not found"**
- Solution: Install Python and add it to PATH

### Runtime Errors

**Error: "Cannot find module './build/Release/Automation.node'"**
- Solution: Run `npm run build` first

**Error: "Access Denied"**
- Solution: Some apps require elevated privileges. Try running as Administrator.

**Element not found**
- Make sure the window is visible and not minimized
- Try using `TreeScopes.Descendants` for a deeper search
- Verify the property name is correct (it may include trailing spaces)

## Next Steps

1. **Explore the tests**: Check out the `tests/` directory for more examples
2. **Read the TypeScript definitions**: Open `index.d.ts` to see all available APIs
3. **Check the official docs**: [Windows UI Automation Documentation](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)
4. **Experiment**: Try automating different applications!

## Tips

- Always check if elements exist before using them
- Use try-catch blocks for error handling
- Add delays between actions if needed (UI updates aren't instant)
- Use Inspect.exe (part of Windows SDK) to explore UI element properties
- Start with simple automations and gradually increase complexity

## Need Help?

- Check the [README.md](README.md) for more information
- Review the example files in the `examples/` directory
- Consult the [Windows UI Automation documentation](https://docs.microsoft.com/en-us/windows/win32/winauto/entry-uiauto-win32)

Happy Automating! 🤖
