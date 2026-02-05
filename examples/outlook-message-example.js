/**
 * Outlook New Message Automation Example
 *
 * This example demonstrates how to:
 * 1. Find an open Outlook message window
 * 2. Fill in recipient, subject, and body
 * 3. Optionally send the message
 *
 * Prerequisites:
 * - Microsoft Outlook CLASSIC desktop app (not the new web-based Outlook)
 * - Create a new message manually (Ctrl+N) before running
 * - Build the addon first: npm run build
 *
 * Note: This example is designed for classic Outlook desktop app.
 * The new web-based Outlook (similar to Teams) uses Chrome-based UI
 * which has limited UI Automation support. For web-based Outlook,
 * consider using Microsoft Graph API instead.
 */

const { Automation, PropertyIds, ControlTypeIds, TreeScopes, PatternIds } = require('..');

async function createOutlookMessage(recipient, subject, body, sendMessage = false) {
  try {
    console.log('Initializing UI Automation...');
    const automation = new Automation();
    const root = automation.getRootElement();

    // Find Outlook window
    console.log('Looking for Outlook window...');
    
    // Try different possible window names
    let outlookWindow = root.findFirst(
      TreeScopes.Children,
      automation.createPropertyCondition(PropertyIds.NamePropertyId, 'Inbox - Outlook')
    );

    if (!outlookWindow) {
      outlookWindow = root.findFirst(
        TreeScopes.Children,
        automation.createPropertyCondition(PropertyIds.NamePropertyId, 'Microsoft Outlook')
      );
    }

    if (!outlookWindow) {
      // Try finding by class name
      const windows = root.findAll(
        TreeScopes.Children,
        automation.createPropertyCondition(PropertyIds.ControlTypePropertyId, ControlTypeIds.WindowControlTypeId)
      );

      for (let i = 0; i < windows.length; i++) {
        const win = windows[i];
        const name = win.currentName;
        if (name && name.includes('Outlook')) {
          outlookWindow = win;
          break;
        }
      }
    }

    if (!outlookWindow) {
      console.log('Outlook window not found. Please open Outlook and try again.');
      return;
    }

    console.log('Found Outlook window:', outlookWindow.currentName);

    // Find the new message window (user should have created it manually)
    console.log('Looking for new message window...');
    console.log('Please make sure you have created a new message (Ctrl+N) before running this script.');
    await new Promise(r => setTimeout(r, 1000));

    let messageWindow = root.findFirst(
      TreeScopes.Children,
      automation.createPropertyCondition(PropertyIds.NamePropertyId, 'Untitled - Message (HTML)')
    );

    if (!messageWindow) {
      // Try finding any new window with "Message" in the title
      const windows = root.findAll(
        TreeScopes.Children,
        automation.createPropertyCondition(PropertyIds.ControlTypePropertyId, ControlTypeIds.WindowControlTypeId)
      );

      for (let i = 0; i < windows.length; i++) {
        const win = windows[i];
        const name = win.currentName;
        if (name && (name.includes('Message') || name.includes('Untitled'))) {
          messageWindow = win;
          break;
        }
      }
    }

    if (!messageWindow) {
      console.log('New message window not found.');
      return;
    }

    console.log('Found message window:', messageWindow.currentName);

    // Fill in To field
    console.log('Filling in recipient...');
    const toField = messageWindow.findFirst(
      TreeScopes.Descendants,
      automation.createPropertyCondition(PropertyIds.AutomationIdPropertyId, '4100')
    );

    if (toField) {
      toField.setFocus();
      await new Promise(r => setTimeout(r, 300));
      const valuePattern = toField.getCurrentPattern(PatternIds.ValuePatternId);
      if (valuePattern) {
        valuePattern.setValue(recipient);
        await new Promise(r => setTimeout(r, 500));
      }
    } else {
      console.log('To field not found by AutomationId, trying by name...');
      const toFieldAlt = messageWindow.findFirst(
        TreeScopes.Descendants,
        automation.createPropertyCondition(PropertyIds.NamePropertyId, 'To')
      );
      if (toFieldAlt) {
        toFieldAlt.setFocus();
        await new Promise(r => setTimeout(r, 300));
        const valuePattern = toFieldAlt.getCurrentPattern(PatternIds.ValuePatternId);
        if (valuePattern) {
          valuePattern.setValue(recipient);
          await new Promise(r => setTimeout(r, 500));
        }
      }
    }

    // Fill in Subject field
    console.log('Filling in subject...');
    const subjectField = messageWindow.findFirst(
      TreeScopes.Descendants,
      automation.createPropertyCondition(PropertyIds.AutomationIdPropertyId, '4101')
    );

    if (subjectField) {
      subjectField.setFocus();
      await new Promise(r => setTimeout(r, 300));
      const valuePattern = subjectField.getCurrentPattern(PatternIds.ValuePatternId);
      if (valuePattern) {
        valuePattern.setValue(subject);
        await new Promise(r => setTimeout(r, 500));
      }
    } else {
      console.log('Subject field not found by AutomationId, trying by name...');
      const subjectFieldAlt = messageWindow.findFirst(
        TreeScopes.Descendants,
        automation.createPropertyCondition(PropertyIds.NamePropertyId, 'Subject')
      );
      if (subjectFieldAlt) {
        subjectFieldAlt.setFocus();
        await new Promise(r => setTimeout(r, 300));
        const valuePattern = subjectFieldAlt.getCurrentPattern(PatternIds.ValuePatternId);
        if (valuePattern) {
          valuePattern.setValue(subject);
          await new Promise(r => setTimeout(r, 500));
        }
      }
    }

    // Fill in message body
    console.log('Filling in message body...');
    
    // Find the document/edit control for the body
    let bodyControl = messageWindow.findFirst(
      TreeScopes.Descendants,
      automation.createPropertyCondition(PropertyIds.ControlTypePropertyId, ControlTypeIds.DocumentControlTypeId)
    );

    if (!bodyControl) {
      // Try finding edit control
      const editControls = messageWindow.findAll(
        TreeScopes.Descendants,
        automation.createPropertyCondition(PropertyIds.ControlTypePropertyId, ControlTypeIds.EditControlTypeId)
      );

      // Find the largest edit control (likely the body)
      for (let i = 0; i < editControls.length; i++) {
        const control = editControls[i];
        const rect = control.currentBoundingRectangle;
        if (rect && rect.bottom - rect.top > 100) { // Large enough to be body
          bodyControl = control;
          break;
        }
      }
    }

    if (bodyControl) {
      console.log('Found body control, setting text...');
      bodyControl.setFocus();
      await new Promise(r => setTimeout(r, 500));
      
      const valuePattern = bodyControl.getCurrentPattern(PatternIds.ValuePatternId);
      if (valuePattern) {
        valuePattern.setValue(body);
        await new Promise(r => setTimeout(r, 500));
      }
    } else {
      console.log('Message body control not found.');
    }

    if (sendMessage) {
      // Find and click Send button
      console.log('Looking for Send button...');
      const sendButton = messageWindow.findFirst(
        TreeScopes.Descendants,
        automation.createPropertyCondition(PropertyIds.NamePropertyId, 'Send')
      );

      if (sendButton) {
        console.log('Clicking Send button...');
        const invokePattern = sendButton.getCurrentPattern(PatternIds.InvokePatternId);
        if (invokePattern) {
          invokePattern.invoke();
          console.log('✓ Email sent successfully!');
        }
      } else {
        console.log('Send button not found.');
      }
    } else {
      console.log('✓ Email draft created successfully!');
      console.log('Message is ready to send. Click Send manually or set sendMessage=true.');
    }

  } catch (error) {
    console.error('Error during automation:', error.message);
    console.error(error);
  }
}

// Run the automation
console.log('=== Outlook New Message Automation Example ===\n');

// Example usage - modify these values
const RECIPIENT = 'example@example.com';
const SUBJECT = 'Test Email from node-winautomation';
const BODY = `Hello,

This email was created using node-winautomation, a Node.js Windows UI Automation library.

Best regards,
Automated System`;

const SEND_MESSAGE = false; // Set to true to actually send the email

createOutlookMessage(RECIPIENT, SUBJECT, BODY, SEND_MESSAGE).then(() => {
  console.log('\nExample completed.');
}).catch(err => {
  console.error('Unexpected error:', err);
  process.exit(1);
});
