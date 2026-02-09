/**
 * Outlook Automation Example
 * 
 * This example demonstrates how to use the Outlook connector to:
 * - Create and send emails
 * - Create appointments
 * - Create contacts
 * - Access folders and items
 * 
 * Prerequisites:
 * - Microsoft Outlook installed
 * - Build the addon first: npm run build
 */

// Standard professional import style
const { OutlookConnector, OlDefaultFolders, OlImportance } = require('../../../index.js');

// Example 1: Create and Send Email
async function createEmail() {
    console.log('=== Example 1: Create Email ===\n');
    
    const outlook = new OutlookConnector();
    
    try {
        // Create a new mail item
        const mail = outlook.createMailItem();
        
        // Set email properties
        mail.setTo('recipient@example.com');
        mail.setCC('cc@example.com');
        mail.setSubject('Hello from node-winautomation!');
        mail.setBody('This email was created using the Outlook connector.\n\nBest regards,\nAutomated System');
        
        // Set importance
        mail.setImportance(OlImportance.olImportanceHigh);
        
        // Save as draft (don't send)
        mail.save();
        console.log('✓ Email saved as draft');
        
        // To send the email, uncomment:
        // mail.send();
        // console.log('✓ Email sent');
        
    } catch (error) {
        console.error('❌ Failed to create email:', error.message);
        throw error;
    } finally {
        outlook.release();
    }
}

// Example 2: Create Appointment
async function createAppointment() {
    console.log('\n=== Example 2: Create Appointment ===\n');
    
    const outlook = new OutlookConnector();
    
    try {
        // Create appointment
        const appointment = outlook.createAppointmentItem();
        
        // Set appointment properties
        appointment.setSubject('Team Meeting');
        appointment.setLocation('Conference Room A');
        appointment.setBody('Quarterly review meeting');
        
        // Set date/time (example: tomorrow at 2 PM)
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        tomorrow.setHours(14, 0, 0, 0);
        appointment.setStart(tomorrow);
        
        const endTime = new Date(tomorrow);
        endTime.setHours(15, 0, 0, 0);
        appointment.setEnd(endTime);
        
        // Set reminder
        appointment.setReminderSet(true);
        appointment.setReminderMinutesBeforeStart(15);
        
        // Save appointment
        appointment.save();
        console.log('✓ Appointment created');
        
    } catch (error) {
        console.error('❌ Failed to create appointment:', error.message);
        throw error;
    } finally {
        outlook.release();
    }
}

// Example 3: Create Contact
async function createContact() {
    console.log('\n=== Example 3: Create Contact ===\n');
    
    const outlook = new OutlookConnector();
    
    try {
        // Create contact
        const contact = outlook.createContactItem();
        
        // Set contact properties
        contact.setFirstName('John');
        contact.setLastName('Doe');
        contact.setEmail1Address('john.doe@example.com');
        contact.setBusinessTelephoneNumber('+1-555-0100');
        contact.setMobileTelephoneNumber('+1-555-0101');
        contact.setCompanyName('Example Corp');
        contact.setJobTitle('Software Engineer');
        
        // Save contact
        contact.save();
        console.log('✓ Contact created');
        
    } catch (error) {
        console.error('❌ Failed to create contact:', error.message);
        throw error;
    } finally {
        outlook.release();
    }
}

// Example 4: Access Inbox and Read Emails
async function readInbox() {
    console.log('\n=== Example 4: Read Inbox ===\n');
    
    const outlook = new OutlookConnector();
    
    try {
        // Get inbox folder
        const inbox = outlook.getDefaultFolder(OlDefaultFolders.olFolderInbox);
        
        console.log(`Folder: ${inbox.getName()}`);
        console.log(`Total items: ${inbox.getCount()}`);
        console.log(`Unread items: ${inbox.getUnReadItemCount()}`);
        
        // Get first item (if any)
        if (inbox.getCount() > 0) {
            const firstItem = inbox.getItem(1);
            console.log('\nFirst email:');
            console.log(`  Subject: ${firstItem.getProperty('Subject')}`);
            console.log(`  From: ${firstItem.getProperty('SenderName')}`);
        }
        
    } catch (error) {
        console.error('❌ Failed to read inbox:', error.message);
        throw error;
    } finally {
        outlook.release();
    }
}

// Example 5: Create Email with HTML Body
async function createHtmlEmail() {
    console.log('\n=== Example 5: Create HTML Email ===\n');
    
    const outlook = new OutlookConnector();
    
    try {
        const mail = outlook.createMailItem();
        
        mail.setTo('recipient@example.com');
        mail.setSubject('HTML Email Example');
        
        // Set HTML body
        const htmlBody = `
            <html>
            <body>
                <h2>Hello from node-winautomation!</h2>
                <p>This is an <strong>HTML email</strong> with formatting:</p>
                <ul>
                    <li>Bullet point 1</li>
                    <li>Bullet point 2</li>
                    <li>Bullet point 3</li>
                </ul>
                <p style="color: blue;">This text is blue!</p>
            </body>
            </html>
        `;
        
        mail.setHTMLBody(htmlBody);
        mail.save();
        
        console.log('✓ HTML email created');
        
    } catch (error) {
        console.error('❌ Failed to create HTML email:', error.message);
        throw error;
    } finally {
        outlook.release();
    }
}

// Example 6: Create Task
async function createTask() {
    console.log('\n=== Example 6: Create Task ===\n');
    
    const outlook = new OutlookConnector();
    
    try {
        const task = outlook.createTaskItem();
        
        task.setSubject('Complete project documentation');
        task.setBody('Write comprehensive documentation for the project');
        
        // Set due date (one week from now)
        const dueDate = new Date();
        dueDate.setDate(dueDate.getDate() + 7);
        task.setDueDate(dueDate);
        
        // Set status
        task.setStatus(1); // olTaskNotStarted
        
        // Set percent complete
        task.setPercentComplete(0);
        
        task.save();
        console.log('✓ Task created');
        
    } catch (error) {
        console.error('❌ Failed to create task:', error.message);
        throw error;
    } finally {
        outlook.release();
    }
}

// Run all examples
async function runExamples() {
    console.log('Outlook Connector Examples\n');
    console.log('=' .repeat(50) + '\n');
    
    try {
        await createEmail();
        await createAppointment();
        await createContact();
        await readInbox();
        await createHtmlEmail();
        await createTask();
        
        console.log('\n' + '='.repeat(50));
        console.log('✅ All examples completed successfully!');
        console.log('\nCheck your Outlook for:');
        console.log('  - Draft emails in Drafts folder');
        console.log('  - New appointment in Calendar');
        console.log('  - New contact in Contacts');
        console.log('  - New task in Tasks');
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        console.error(error.stack);
    }
}

// Run examples
runExamples();
