# Microsoft Outlook COM API Research

## Overview

Microsoft Outlook exposes a rich COM API through the `Outlook.Application` object. The API follows a hierarchical object model.

## Object Model Hierarchy

```
Outlook.Application
├── Session (NameSpace)
│   ├── Folders (collection)
│   ├── CurrentUser
│   ├── Accounts (collection)
│   └── Stores (collection)
├── Explorers (collection)
├── Inspectors (collection)
├── ActiveExplorer
├── ActiveInspector
└── CreateItem(ItemType)
    ├── MailItem
    ├── AppointmentItem
    ├── ContactItem
    ├── TaskItem
    └── MeetingItem
```

## Core Interfaces

### 1. Application Object
**ProgID:** `Outlook.Application`

**Key Properties:**
- `Version` (string) - Outlook version
- `Name` (string) - Application name
- `Session` (NameSpace) - Current session
- `ActiveExplorer` (Explorer) - Active explorer window
- `ActiveInspector` (Inspector) - Active inspector window

**Key Methods:**
- `CreateItem(ItemType)` - Create new Outlook item
- `GetNamespace(Type)` - Get namespace object (usually "MAPI")
- `Quit()` - Close Outlook

**Events:**
- `Startup` - Fired when Outlook starts
- `Quit` - Fired when Outlook quits
- `ItemSend(Item, Cancel)` - Fired before item is sent
- `NewMailEx(EntryIDCollection)` - Fired when new mail arrives

### 2. MailItem Object
**Created via:** `Application.CreateItem(0)` where 0 = olMailItem

**Key Properties:**
- `Subject` (string) - Email subject
- `Body` (string) - Plain text body
- `HTMLBody` (string) - HTML body
- `To` (string) - Recipients (semicolon-separated)
- `CC` (string) - CC recipients
- `BCC` (string) - BCC recipients
- `From` (string) - Sender address
- `SentOn` (Date) - Sent date/time
- `ReceivedTime` (Date) - Received date/time
- `Importance` (OlImportance) - Importance level (0=Low, 1=Normal, 2=High)
- `Sensitivity` (OlSensitivity) - Sensitivity level
- `Attachments` (Attachments) - Attachment collection
- `Recipients` (Recipients) - Recipient collection
- `Categories` (string) - Categories
- `FlagStatus` (OlFlagStatus) - Flag status
- `UnRead` (boolean) - Unread status
- `Size` (long) - Message size in bytes

**Key Methods:**
- `Send()` - Send the email
- `Display(Modal)` - Display the email window
- `Save()` - Save to Drafts
- `Close(SaveMode)` - Close the item
- `Copy()` - Create a copy
- `Move(DestFolder)` - Move to folder
- `Reply()` - Create reply
- `ReplyAll()` - Create reply all
- `Forward()` - Create forward

### 3. AppointmentItem Object
**Created via:** `Application.CreateItem(1)` where 1 = olAppointmentItem

**Key Properties:**
- `Subject` (string) - Appointment subject
- `Location` (string) - Location
- `Start` (Date) - Start date/time
- `End` (Date) - End date/time
- `AllDayEvent` (boolean) - All-day event flag
- `Body` (string) - Appointment body
- `ReminderSet` (boolean) - Reminder enabled
- `ReminderMinutesBeforeStart` (long) - Reminder time
- `BusyStatus` (OlBusyStatus) - Free/Busy status
- `Recipients` (Recipients) - Attendees
- `Organizer` (string) - Meeting organizer
- `IsRecurring` (boolean) - Recurring flag
- `RecurrencePattern` (RecurrencePattern) - Recurrence settings

**Key Methods:**
- `Send()` - Send meeting invitation
- `Display(Modal)` - Display appointment window
- `Save()` - Save appointment
- `GetRecurrencePattern()` - Get recurrence pattern

### 4. ContactItem Object
**Created via:** `Application.CreateItem(2)` where 2 = olContactItem

**Key Properties:**
- `FirstName` (string)
- `LastName` (string)
- `FullName` (string)
- `Email1Address` (string)
- `Email2Address` (string)
- `Email3Address` (string)
- `BusinessTelephoneNumber` (string)
- `HomeTelephoneNumber` (string)
- `MobileTelephoneNumber` (string)
- `CompanyName` (string)
- `JobTitle` (string)
- `BusinessAddress` (string)
- `HomeAddress` (string)
- `WebPage` (string)
- `Birthday` (Date)
- `Categories` (string)

**Key Methods:**
- `Save()` - Save contact
- `Display(Modal)` - Display contact window
- `Delete()` - Delete contact

### 5. TaskItem Object
**Created via:** `Application.CreateItem(3)` where 3 = olTaskItem

**Key Properties:**
- `Subject` (string)
- `Body` (string)
- `StartDate` (Date)
- `DueDate` (Date)
- `Status` (OlTaskStatus)
- `Priority` (OlImportance)
- `PercentComplete` (long)
- `Complete` (boolean)
- `Owner` (string)
- `Categories` (string)

**Key Methods:**
- `Save()` - Save task
- `Display(Modal)` - Display task window
- `Assign()` - Assign task
- `Send()` - Send task request

### 6. Folder Object

**Key Properties:**
- `Name` (string) - Folder name
- `FolderPath` (string) - Full folder path
- `Items` (Items) - Items collection
- `Folders` (Folders) - Subfolders collection
- `DefaultItemType` (OlItemType) - Default item type
- `UnReadItemCount` (long) - Unread count

**Key Methods:**
- `Delete()` - Delete folder
- `CopyTo(DestFolder)` - Copy folder
- `MoveTo(DestFolder)` - Move folder

### 7. NameSpace Object (Session)
**Obtained via:** `Application.GetNamespace("MAPI")`

**Key Properties:**
- `CurrentUser` (Recipient) - Current user
- `Folders` (Folders) - Root folders
- `Accounts` (Accounts) - Email accounts
- `DefaultStore` (Store) - Default store

**Key Methods:**
- `GetDefaultFolder(FolderType)` - Get default folder (Inbox, Sent, etc.)
- `GetFolderFromID(EntryID)` - Get folder by ID
- `GetItemFromID(EntryID)` - Get item by ID
- `Logon(Profile, Password, ShowDialog, NewSession)` - Logon to profile
- `Logoff()` - Logoff

**Default Folders (OlDefaultFolders):**
- `olFolderInbox` = 6
- `olFolderSentMail` = 5
- `olFolderOutbox` = 4
- `olFolderDeletedItems` = 3
- `olFolderDrafts` = 16
- `olFolderCalendar` = 9
- `olFolderContacts` = 10
- `olFolderTasks` = 13

### 8. Attachments Collection

**Key Methods:**
- `Add(Source, Type, Position, DisplayName)` - Add attachment
- `Remove(Index)` - Remove attachment
- `Item(Index)` - Get attachment by index

**Attachment Types:**
- `olByValue` = 1 - File attachment
- `olByReference` = 4 - Shortcut
- `olEmbeddeditem` = 5 - Embedded Outlook item
- `olOLE` = 6 - OLE object

### 9. Recipients Collection

**Key Methods:**
- `Add(Name)` - Add recipient
- `Remove(Index)` - Remove recipient
- `ResolveAll()` - Resolve all recipients

**Recipient Object Properties:**
- `Name` (string)
- `Address` (string)
- `Type` (OlMailRecipientType) - To=1, CC=2, BCC=3
- `Resolved` (boolean)

## Event Interfaces

### ApplicationEvents_11 Interface
**IID:** `{0006304E-0000-0000-C000-000000000046}`

**Events:**
- `ItemSend(Item, Cancel)` - DispID: 61445
- `NewMailEx(EntryIDCollection)` - DispID: 64181
- `Reminder(Item)` - DispID: 62034
- `OptionsPagesAdd(Pages)` - DispID: 61443
- `Startup()` - DispID: 61446
- `Quit()` - DispID: 61447
- `AdvancedSearchComplete(SearchObject)` - DispID: 64180
- `AdvancedSearchStopped(SearchObject)` - DispID: 64179

### ItemEvents_10 Interface (for MailItem, etc.)
**IID:** `{0006302B-0000-0000-C000-000000000046}`

**Events:**
- `Open(Cancel)` - DispID: 61443
- `Close(Cancel)` - DispID: 61448
- `Read()` - DispID: 61441
- `Write(Cancel)` - DispID: 61450
- `Send(Cancel)` - DispID: 61445
- `Reply(Response, Cancel)` - DispID: 61446
- `ReplyAll(Response, Cancel)` - DispID: 61447
- `Forward(Forward, Cancel)` - DispID: 61448
- `AttachmentAdd(Attachment)` - DispID: 64167
- `AttachmentRead(Attachment)` - DispID: 64168
- `AttachmentRemove(Attachment)` - DispID: 64169

## Enumerations

### OlItemType
- `olMailItem` = 0
- `olAppointmentItem` = 1
- `olContactItem` = 2
- `olTaskItem` = 3
- `olJournalItem` = 4
- `olNoteItem` = 5
- `olPostItem` = 6
- `olDistributionListItem` = 7

### OlImportance
- `olImportanceLow` = 0
- `olImportanceNormal` = 1
- `olImportanceHigh` = 2

### OlBusyStatus
- `olFree` = 0
- `olTentative` = 1
- `olBusy` = 2
- `olOutOfOffice` = 3
- `olWorkingElsewhere` = 4

### OlSaveAsType
- `olTXT` = 0
- `olRTF` = 1
- `olTemplate` = 2
- `olMSG` = 3
- `olDoc` = 4
- `olHTML` = 5
- `olVCard` = 6
- `olVCal` = 7
- `olICal` = 8

## Common Patterns

### Creating and Sending Email
```javascript
const app = new OutlookConnector();
const mail = app.createMailItem();
mail.setTo('recipient@example.com');
mail.setSubject('Test Email');
mail.setBody('Email body');
mail.send();
```

### Reading Inbox
```javascript
const inbox = app.getDefaultFolder(6); // olFolderInbox
const items = inbox.getItems();
for (let i = 1; i <= items.count; i++) {
    const mail = items.item(i);
    console.log(mail.getSubject());
}
```

### Creating Appointment
```javascript
const appt = app.createAppointmentItem();
appt.setSubject('Meeting');
appt.setLocation('Conference Room');
appt.setStart(new Date('2024-01-15 10:00'));
appt.setEnd(new Date('2024-01-15 11:00'));
appt.save();
```

### Subscribing to Events
```javascript
app.onNewMail((entryId) => {
    console.log('New mail:', entryId);
});

app.onItemSend((item, cancel) => {
    console.log('Sending:', item.getSubject());
});
```

## References

- [Outlook Object Model Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook/object-model)
- [Outlook Primary Interop Assembly](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook)
- [Outlook Event IDs](https://docs.microsoft.com/en-us/office/vba/api/outlook.application.itemevent)
