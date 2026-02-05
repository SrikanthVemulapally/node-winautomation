# Error Handling Guide

## Overview

This guide covers error handling patterns for COM automation in node-winautomation. Proper error handling is critical for robust COM automation.

## Error Types

### 1. COM Initialization Errors

**Cause:** COM not initialized or initialization failed

**Example:**
```javascript
try {
    const outlook = new COMObject('Outlook.Application');
} catch (error) {
    if (error.message.includes('Failed to get CLSID')) {
        console.error('Outlook is not installed or not registered');
    } else if (error.message.includes('Failed to create COM object')) {
        console.error('Cannot start Outlook - check permissions');
    }
}
```

### 2. Property Access Errors

**Cause:** Property doesn't exist or access denied

**Example:**
```javascript
try {
    const value = comObject.getProperty('NonExistentProperty');
} catch (error) {
    console.error('Property access failed:', error.message);
    // Fallback to default value
    const value = null;
}
```

### 3. Method Invocation Errors

**Cause:** Method doesn't exist, wrong parameters, or execution failed

**Example:**
```javascript
try {
    const result = comObject.invoke('MethodName', arg1, arg2);
} catch (error) {
    if (error.message.includes('Method not found')) {
        console.error('Method does not exist on this object');
    } else if (error.message.includes('Wrong number of arguments')) {
        console.error('Incorrect parameter count');
    } else if (error.message.includes('Type mismatch')) {
        console.error('Parameter type is incorrect');
    }
}
```

### 4. Event Subscription Errors

**Cause:** Object doesn't support events or invalid IID

**Example:**
```javascript
try {
    outlook.adviseEvent(iid, callback);
} catch (error) {
    if (error.message.includes('does not support connection points')) {
        console.error('This object does not support events');
    } else if (error.message.includes('Connection point not found')) {
        console.error('Invalid event interface IID');
    }
}
```

## Error Handling Patterns

### Pattern 1: Try-Catch with Cleanup

```javascript
const outlook = new OutlookConnector();
try {
    const mail = outlook.createMailItem();
    try {
        mail.setTo('user@example.com');
        mail.setSubject('Test');
        mail.send();
    } finally {
        mail.release(); // Always cleanup
    }
} catch (error) {
    console.error('Email operation failed:', error.message);
} finally {
    outlook.release(); // Always cleanup
}
```

### Pattern 2: Defensive Programming

```javascript
function safeGetProperty(obj, propName, defaultValue = null) {
    try {
        return obj.getProperty(propName);
    } catch (error) {
        console.warn(`Failed to get ${propName}:`, error.message);
        return defaultValue;
    }
}

// Usage
const subject = safeGetProperty(mailItem, 'Subject', '(No Subject)');
```

### Pattern 3: Retry Logic

```javascript
async function sendEmailWithRetry(mail, maxRetries = 3) {
    for (let i = 0; i < maxRetries; i++) {
        try {
            mail.send();
            return true;
        } catch (error) {
            console.warn(`Send attempt ${i + 1} failed:`, error.message);
            if (i === maxRetries - 1) {
                throw error; // Last attempt failed
            }
            await new Promise(r => setTimeout(r, 1000 * (i + 1))); // Exponential backoff
        }
    }
    return false;
}
```

### Pattern 4: Validation Before Operation

```javascript
function createEmail(outlook, to, subject, body) {
    // Validate inputs
    if (!to || typeof to !== 'string') {
        throw new Error('Invalid recipient email address');
    }
    if (!subject || typeof subject !== 'string') {
        throw new Error('Invalid subject');
    }
    
    try {
        const mail = outlook.createMailItem();
        mail.setTo(to);
        mail.setSubject(subject);
        mail.setBody(body || '');
        return mail;
    } catch (error) {
        throw new Error(`Failed to create email: ${error.message}`);
    }
}
```

### Pattern 5: Resource Pool

```javascript
class ExcelPool {
    constructor(maxInstances = 3) {
        this.pool = [];
        this.maxInstances = maxInstances;
    }
    
    acquire() {
        if (this.pool.length > 0) {
            return this.pool.pop();
        }
        if (this.pool.length < this.maxInstances) {
            try {
                const excel = new ExcelConnector();
                excel.setVisible(false);
                excel.setDisplayAlerts(false);
                return excel;
            } catch (error) {
                throw new Error(`Failed to create Excel instance: ${error.message}`);
            }
        }
        throw new Error('Excel pool exhausted');
    }
    
    release(excel) {
        if (this.pool.length < this.maxInstances) {
            this.pool.push(excel);
        } else {
            excel.quit();
            excel.release();
        }
    }
    
    cleanup() {
        for (const excel of this.pool) {
            try {
                excel.quit();
                excel.release();
            } catch (error) {
                console.error('Cleanup error:', error.message);
            }
        }
        this.pool = [];
    }
}
```

## Connector-Specific Error Handling

### OutlookConnector Errors

```javascript
const outlook = new OutlookConnector();

try {
    // Check if Outlook is available
    const version = outlook.getVersion();
    console.log('Outlook version:', version);
    
    // Create email with validation
    const mail = outlook.createMailItem();
    
    // Validate recipient before sending
    mail.setTo('user@example.com');
    const recipients = mail.getProperty('Recipients');
    recipients.invoke('ResolveAll');
    
    // Check if all recipients resolved
    for (let i = 1; i <= recipients.getProperty('Count'); i++) {
        const recipient = recipients.invoke('Item', i);
        if (!recipient.getProperty('Resolved')) {
            throw new Error(`Recipient ${recipient.getProperty('Name')} could not be resolved`);
        }
    }
    
    mail.send();
    
} catch (error) {
    console.error('Outlook operation failed:', error.message);
    
    // Log detailed error information
    if (error.stack) {
        console.error('Stack trace:', error.stack);
    }
} finally {
    outlook.release();
}
```

### ExcelConnector Errors

```javascript
const excel = new ExcelConnector();

try {
    excel.setVisible(true);
    excel.setDisplayAlerts(false);
    
    // Try to open workbook with fallback
    let workbook;
    try {
        workbook = excel.openWorkbook('C:\\data\\input.xlsx');
    } catch (error) {
        console.warn('Failed to open existing workbook, creating new:', error.message);
        workbook = excel.addWorkbook();
    }
    
    try {
        const sheet = workbook.getActiveSheet();
        
        // Safe cell access
        try {
            sheet.getCells(1, 1).setValue('Data');
        } catch (error) {
            console.error('Failed to write to cell:', error.message);
        }
        
        // Safe save with error handling
        try {
            workbook.save();
        } catch (error) {
            console.error('Save failed, trying SaveAs:', error.message);
            workbook.saveAs('C:\\data\\backup.xlsx');
        }
        
    } finally {
        workbook.close(false); // Don't save on error
    }
    
} catch (error) {
    console.error('Excel operation failed:', error.message);
} finally {
    try {
        excel.quit();
    } catch (error) {
        console.error('Failed to quit Excel:', error.message);
    }
    excel.release();
}
```

## Common Error Messages

### HRESULT Error Codes

| HRESULT | Constant | Meaning | Solution |
|---------|----------|---------|----------|
| 0x80004005 | E_FAIL | Unspecified error | Check COM object state |
| 0x80070005 | E_ACCESSDENIED | Access denied | Run as Administrator |
| 0x80040154 | REGDB_E_CLASSNOTREG | Class not registered | Install/repair Office |
| 0x800401F3 | CO_E_NOTINITIALIZED | COM not initialized | Check COMLifecycle |
| 0x80020006 | DISP_E_UNKNOWNNAME | Unknown name | Property/method doesn't exist |
| 0x8002000E | DISP_E_BADPARAMCOUNT | Bad parameter count | Wrong number of arguments |
| 0x80020005 | DISP_E_TYPEMISMATCH | Type mismatch | Wrong parameter type |

### Error Message Patterns

```javascript
function parseComError(error) {
    const message = error.message;
    
    if (message.includes('HRESULT: 0x')) {
        const match = message.match(/HRESULT: (0x[0-9A-F]+)/i);
        if (match) {
            const hresult = match[1];
            return {
                type: 'COM_ERROR',
                hresult: hresult,
                description: message
            };
        }
    }
    
    if (message.includes('Property not found')) {
        return { type: 'PROPERTY_ERROR', property: extractPropertyName(message) };
    }
    
    if (message.includes('Method not found')) {
        return { type: 'METHOD_ERROR', method: extractMethodName(message) };
    }
    
    return { type: 'UNKNOWN_ERROR', message: message };
}
```

## Best Practices

### 1. Always Use Try-Catch

```javascript
// ✓ Good
try {
    const value = obj.getProperty('Name');
} catch (error) {
    console.error('Error:', error.message);
}

// ✗ Bad
const value = obj.getProperty('Name'); // May throw
```

### 2. Always Cleanup Resources

```javascript
// ✓ Good
const connector = new OutlookConnector();
try {
    // ... operations ...
} finally {
    connector.release(); // Always called
}

// ✗ Bad
const connector = new OutlookConnector();
// ... operations ...
connector.release(); // May not be called if error occurs
```

### 3. Validate Inputs

```javascript
// ✓ Good
function setEmailRecipient(mail, email) {
    if (!email || !email.includes('@')) {
        throw new Error('Invalid email address');
    }
    mail.setTo(email);
}

// ✗ Bad
function setEmailRecipient(mail, email) {
    mail.setTo(email); // May fail with cryptic COM error
}
```

### 4. Provide Context in Errors

```javascript
// ✓ Good
try {
    mail.send();
} catch (error) {
    throw new Error(`Failed to send email to ${mail.getTo()}: ${error.message}`);
}

// ✗ Bad
try {
    mail.send();
} catch (error) {
    throw error; // No context
}
```

### 5. Log Errors Appropriately

```javascript
// ✓ Good
try {
    // ... operation ...
} catch (error) {
    console.error('Operation failed:', {
        message: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString(),
        context: { /* relevant data */ }
    });
}

// ✗ Bad
try {
    // ... operation ...
} catch (error) {
    console.log(error); // Insufficient information
}
```

## Error Recovery Strategies

### Strategy 1: Graceful Degradation

```javascript
function getEmailSafely(outlook, index) {
    try {
        const inbox = outlook.getInbox();
        const item = inbox.getItem(index);
        return {
            subject: safeGet(item, 'Subject', '(No Subject)'),
            from: safeGet(item, 'SenderName', '(Unknown)'),
            date: safeGet(item, 'ReceivedTime', new Date())
        };
    } catch (error) {
        console.error('Failed to get email:', error.message);
        return null; // Graceful failure
    }
}
```

### Strategy 2: Fallback Options

```javascript
function openWorkbook(excel, primaryPath, backupPath) {
    try {
        return excel.openWorkbook(primaryPath);
    } catch (error) {
        console.warn('Primary path failed, trying backup:', error.message);
        try {
            return excel.openWorkbook(backupPath);
        } catch (backupError) {
            console.error('Both paths failed, creating new workbook');
            return excel.addWorkbook();
        }
    }
}
```

### Strategy 3: Transaction-like Operations

```javascript
async function updateContactSafely(outlook, contactId, updates) {
    const contact = outlook.getContactItem(contactId);
    const originalData = {
        firstName: contact.getFirstName(),
        lastName: contact.getLastName(),
        email: contact.getEmail1Address()
    };
    
    try {
        // Apply updates
        if (updates.firstName) contact.setFirstName(updates.firstName);
        if (updates.lastName) contact.setLastName(updates.lastName);
        if (updates.email) contact.setEmail1Address(updates.email);
        
        contact.save();
        return true;
    } catch (error) {
        console.error('Update failed, rolling back:', error.message);
        
        // Rollback
        try {
            contact.setFirstName(originalData.firstName);
            contact.setLastName(originalData.lastName);
            contact.setEmail1Address(originalData.email);
            contact.save();
        } catch (rollbackError) {
            console.error('Rollback failed:', rollbackError.message);
        }
        
        return false;
    }
}
```

## Testing Error Handling

```javascript
describe('Error Handling', () => {
    it('should handle invalid ProgID', () => {
        expect(() => {
            new COMObject('Invalid.ProgID');
        }).toThrow(/Failed to get CLSID/);
    });
    
    it('should handle missing property', () => {
        const outlook = new OutlookConnector();
        const mail = outlook.createMailItem();
        
        expect(() => {
            mail.getProperty('NonExistentProperty');
        }).toThrow(/Property not found/);
    });
    
    it('should cleanup on error', () => {
        const excel = new ExcelConnector();
        let cleaned = false;
        
        try {
            // ... operations that may fail ...
        } finally {
            excel.release();
            cleaned = true;
        }
        
        expect(cleaned).toBe(true);
    });
});
```

## See Also

- [COM API Reference](./COM_API.md)
- [Connectors Guide](./CONNECTORS.md)
- [Memory Management Guide](./MEMORY_MANAGEMENT.md)
