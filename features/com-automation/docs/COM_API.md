# COM Automation API Reference

## Overview

The COM automation system provides a robust, memory-safe way to interact with COM objects from Node.js. It supports property access, method invocation, and event subscriptions.

## Table of Contents

- [Installation](#installation)
- [Basic Usage](#basic-usage)
- [API Reference](#api-reference)
- [Memory Management](#memory-management)
- [Error Handling](#error-handling)
- [Thread Safety](#thread-safety)
- [Best Practices](#best-practices)

## Installation

```bash
npm install node-winautomation
```

## Basic Usage

### Creating COM Objects

```javascript
const { COMObject } = require('node-winautomation');

// Create an Outlook Application object
const outlook = new COMObject('Outlook.Application');

// Create an Excel Application object
const excel = new COMObject('Excel.Application');
```

### Property Access

```javascript
// Get a property
const version = outlook.getProperty('Version');
console.log('Outlook version:', version);

// Set a property
excel.setProperty('Visible', true);
```

### Method Invocation

```javascript
// Invoke a method with no arguments
const workbooks = excel.getProperty('Workbooks');
const workbook = workbooks.invoke('Add');

// Invoke a method with arguments
const mailItem = outlook.invoke('CreateItem', 0); // 0 = olMailItem
mailItem.setProperty('To', 'recipient@example.com');
mailItem.setProperty('Subject', 'Test Email');
mailItem.setProperty('Body', 'This is a test email.');
mailItem.invoke('Send');
```

### Event Subscriptions

```javascript
// Subscribe to events
const eventsIID = '{0006304E-0000-0000-C000-000000000046}'; // Outlook events
outlook.adviseEvent(eventsIID, (event) => {
    console.log('Event fired:', event.dispId);
    console.log('Arguments:', event.args);
});

// Unsubscribe from events
outlook.unadviseEvent(eventsIID);
```

## API Reference

### COMObject Class

#### Constructor

```javascript
new COMObject(progId: string)
```

Creates a new COM object from a ProgID.

**Parameters:**
- `progId` (string): The programmatic identifier (e.g., 'Outlook.Application')

**Throws:**
- Error if the ProgID is invalid
- Error if COM object creation fails

**Example:**
```javascript
const outlook = new COMObject('Outlook.Application');
```

#### getProperty(name)

```javascript
getProperty(name: string): any
```

Gets the value of a property.

**Parameters:**
- `name` (string): Property name

**Returns:**
- Property value (type depends on the property)

**Throws:**
- Error if property doesn't exist
- Error if property access fails

**Example:**
```javascript
const version = outlook.getProperty('Version');
const visible = excel.getProperty('Visible');
```

#### setProperty(name, value)

```javascript
setProperty(name: string, value: any): void
```

Sets the value of a property.

**Parameters:**
- `name` (string): Property name
- `value` (any): Value to set

**Throws:**
- Error if property doesn't exist
- Error if property is read-only
- Error if type conversion fails

**Example:**
```javascript
excel.setProperty('Visible', true);
mailItem.setProperty('Subject', 'Hello World');
```

#### invoke(method, ...args)

```javascript
invoke(method: string, ...args: any[]): any
```

Invokes a method on the COM object.

**Parameters:**
- `method` (string): Method name
- `...args` (any[]): Method arguments

**Returns:**
- Method return value (type depends on the method)

**Throws:**
- Error if method doesn't exist
- Error if argument types are invalid
- Error if method invocation fails

**Example:**
```javascript
const mailItem = outlook.invoke('CreateItem', 0);
const workbook = workbooks.invoke('Open', 'C:\\data.xlsx');
mailItem.invoke('Send');
```

#### adviseEvent(iid, callback)

```javascript
adviseEvent(iid: string, callback: (event: EventData) => void): number
```

Subscribes to COM events.

**Parameters:**
- `iid` (string): Interface ID of the event interface (GUID format)
- `callback` (function): Event handler function

**Returns:**
- Cookie value for the event subscription

**Throws:**
- Error if IID is invalid
- Error if object doesn't support events
- Error if connection point not found

**Example:**
```javascript
const cookie = outlook.adviseEvent(
    '{0006304E-0000-0000-C000-000000000046}',
    (event) => {
        if (event.dispId === 15) { // NewMailEx
            console.log('New mail arrived:', event.args[0]);
        }
    }
);
```

#### unadviseEvent(iid)

```javascript
unadviseEvent(iid: string): void
```

Unsubscribes from COM events.

**Parameters:**
- `iid` (string): Interface ID of the event interface

**Throws:**
- Error if IID is invalid
- Error if no subscription exists

**Example:**
```javascript
outlook.unadviseEvent('{0006304E-0000-0000-C000-000000000046}');
```

#### release()

```javascript
release(): void
```

Manually releases the COM object. This is called automatically when the object is garbage collected, but can be called explicitly for immediate cleanup.

**Example:**
```javascript
outlook.release();
```

### EventData Interface

```typescript
interface EventData {
    dispId: number;      // Dispatch ID of the event
    args: any[];         // Event arguments
}
```

## Memory Management

### Automatic Cleanup

COM objects are automatically cleaned up when they are garbage collected:

```javascript
function createOutlook() {
    const outlook = new COMObject('Outlook.Application');
    // ... use outlook ...
}  // outlook is automatically released when function exits
```

### Manual Cleanup

For long-running applications, you can manually release objects:

```javascript
const outlook = new COMObject('Outlook.Application');
try {
    // ... use outlook ...
} finally {
    outlook.release();  // Explicit cleanup
}
```

### Event Cleanup

Event subscriptions are automatically cleaned up when the COM object is released:

```javascript
const outlook = new COMObject('Outlook.Application');
outlook.adviseEvent(iid, callback);
// ... use outlook ...
outlook.release();  // Events are automatically unadvised
```

## Error Handling

### Try-Catch Pattern

Always wrap COM operations in try-catch blocks:

```javascript
try {
    const outlook = new COMObject('Outlook.Application');
    const version = outlook.getProperty('Version');
    console.log('Version:', version);
} catch (error) {
    console.error('COM error:', error.message);
}
```

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Failed to get CLSID from ProgID" | Invalid ProgID | Check ProgID spelling |
| "Failed to create COM object" | Application not installed | Install the application |
| "Property not found" | Invalid property name | Check documentation |
| "Method not found" | Invalid method name | Check documentation |
| "Connection point not found" | Invalid event IID | Check event interface GUID |

## Thread Safety

### Main Thread

COM objects should be created and used on the main Node.js thread:

```javascript
// ✓ Good - main thread
const outlook = new COMObject('Outlook.Application');

// ✗ Bad - worker thread (may cause issues)
const { Worker } = require('worker_threads');
const worker = new Worker('./worker.js');
```

### Event Callbacks

Event callbacks are automatically marshaled to the main thread and are safe to use:

```javascript
outlook.adviseEvent(iid, (event) => {
    // This callback runs on the main thread
    // Safe to access Node.js APIs and V8 objects
    console.log('Event:', event.dispId);
});
```

## Best Practices

### 1. Use Try-Finally for Cleanup

```javascript
const outlook = new COMObject('Outlook.Application');
try {
    // ... use outlook ...
} finally {
    outlook.release();
}
```

### 2. Check for Null/Undefined

```javascript
const item = outlook.invoke('GetItemFromID', id);
if (item) {
    const subject = item.getProperty('Subject');
    console.log(subject);
    item.release();
}
```

### 3. Unsubscribe from Events

```javascript
const iid = '{0006304E-0000-0000-C000-000000000046}';
outlook.adviseEvent(iid, callback);

// When done:
outlook.unadviseEvent(iid);
```

### 4. Handle Errors Gracefully

```javascript
try {
    mailItem.invoke('Send');
} catch (error) {
    console.error('Failed to send email:', error.message);
    // Handle error appropriately
}
```

### 5. Release Nested Objects

```javascript
const workbooks = excel.getProperty('Workbooks');
const workbook = workbooks.invoke('Add');

// ... use workbook ...

workbook.release();
workbooks.release();
excel.release();
```

### 6. Use Constants

```javascript
// Define constants for magic numbers
const olMailItem = 0;
const olAppointmentItem = 1;

const mailItem = outlook.invoke('CreateItem', olMailItem);
```

### 7. Validate Input

```javascript
function sendEmail(to, subject, body) {
    if (!to || !subject || !body) {
        throw new Error('Missing required parameters');
    }
    
    const outlook = new COMObject('Outlook.Application');
    try {
        const mail = outlook.invoke('CreateItem', 0);
        mail.setProperty('To', to);
        mail.setProperty('Subject', subject);
        mail.setProperty('Body', body);
        mail.invoke('Send');
    } finally {
        outlook.release();
    }
}
```

## Performance Tips

### 1. Reuse COM Objects

```javascript
// ✓ Good - reuse
const outlook = new COMObject('Outlook.Application');
for (let i = 0; i < 100; i++) {
    const mail = outlook.invoke('CreateItem', 0);
    // ... configure mail ...
    mail.invoke('Send');
}
outlook.release();

// ✗ Bad - recreate
for (let i = 0; i < 100; i++) {
    const outlook = new COMObject('Outlook.Application');
    // ... use outlook ...
    outlook.release();
}
```

### 2. Batch Operations

```javascript
// Get multiple properties at once if possible
const props = ['Subject', 'Body', 'To', 'From'];
const values = props.map(prop => mailItem.getProperty(prop));
```

### 3. Minimize Cross-Process Calls

```javascript
// ✓ Good - fewer calls
const workbook = workbooks.invoke('Add');
const sheets = workbook.getProperty('Worksheets');
const sheet = sheets.getProperty('Item', 1);

// ✗ Bad - more calls
const sheet = excel
    .getProperty('Workbooks')
    .invoke('Add')
    .getProperty('Worksheets')
    .getProperty('Item', 1);
```

## Examples

See the [examples/com](../examples/com) directory for complete working examples:

- `outlook-basic.js` - Basic Outlook automation
- `outlook-events.js` - Event handling
- `excel-basic.js` - Excel automation
- `word-basic.js` - Word automation

## Troubleshooting

### COM Object Not Found

**Problem:** "Failed to get CLSID from ProgID"

**Solution:**
- Verify the application is installed
- Check the ProgID spelling
- Run `reg query HKCR\<ProgID>` to verify registration

### Access Denied

**Problem:** "Access is denied"

**Solution:**
- Run as Administrator
- Check DCOM permissions
- Verify application security settings

### Events Not Firing

**Problem:** Event callback never called

**Solution:**
- Verify correct event IID
- Ensure message pump is running (Node.js event loop)
- Check event is actually being triggered

## See Also

- [Connectors Guide](./CONNECTORS.md) - Pre-built connectors for common applications
- [Memory Management](./MEMORY_MANAGEMENT.md) - Detailed memory management guide
- [Architecture](../ARCHITECTURE.md) - System architecture overview
