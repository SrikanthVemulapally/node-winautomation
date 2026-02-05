# COM Automation Examples

This folder contains practical examples demonstrating how to use the COM Automation connectors for Microsoft Office applications.

## Examples

### Outlook Connector (`outlook-example.js`)

Demonstrates Outlook automation including:
- Creating and sending emails
- Creating appointments and calendar events
- Creating contacts
- Creating tasks
- Accessing folders (Inbox, Drafts, etc.)
- Working with HTML email bodies
- Setting email importance and properties

**Run:**
```bash
node features/com-automation/examples/outlook-example.js
```

### Excel Connector (`excel-example.js`)

Demonstrates Excel automation including:
- Creating workbooks and worksheets
- Working with multiple sheets
- Setting calculation modes
- Controlling display settings (visibility, alerts, screen updating)
- Saving in different formats (XLSX, CSV, HTML)
- Opening existing workbooks
- Getting Excel version information

**Run:**
```bash
node features/com-automation/examples/excel-example.js
```

## Prerequisites

- Microsoft Outlook and/or Excel installed
- Build the native addon: `npm run build`

## Usage

Each example file can be run independently:

```bash
# Run Outlook examples
node features/com-automation/examples/outlook-example.js

# Run Excel examples
node features/com-automation/examples/excel-example.js
```

## Import Style

All examples use the recommended namespace import style:

```javascript
const { COMAutomation } = require('../../../index.js');

// Create connectors
const outlook = new COMAutomation.OutlookConnector();
const excel = new COMAutomation.ExcelConnector();
```

## Notes

- **Outlook Examples**: Creates draft items by default. To actually send emails, uncomment the `mail.send()` line.
- **Excel Examples**: Cell operations (Range/Cells) require parameterized property support. Current examples focus on workbook and worksheet operations.
- All examples include proper resource cleanup with `release()` calls.

## API Reference

For complete API documentation, see:
- `features/com-automation/docs/` - Detailed connector documentation
- `README.md` - Main project documentation
- `index.d.ts` - TypeScript definitions

## Example Output

### Outlook Example
```
✓ Email saved as draft
✓ Appointment created
✓ Contact created
✓ Task created
```

### Excel Example
```
✓ Workbook created
✓ Sheet renamed
✓ Workbook saved
✓ Multiple sheets created
✓ Files saved in various formats
```

## Customization

Feel free to modify these examples for your specific use cases:
- Change email recipients and content
- Adjust appointment times and locations
- Modify contact information
- Customize Excel file paths and formats

## Troubleshooting

**"Cannot find module"**: Make sure you've built the native addon first:
```bash
npm run build
```

**"Outlook/Excel not found"**: Ensure Microsoft Office is installed and properly registered.

**"Access Denied"**: Some operations may require elevated privileges. Try running as Administrator.

## Additional Resources

- [Outlook Object Model Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook)
- [Excel Object Model Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
