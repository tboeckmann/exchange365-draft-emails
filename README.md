# Exchange365 Draft Emails

A Node.js library for creating draft emails in Exchange/Office 365 using Microsoft Graph API.

## Setup

1. Register an application in Azure AD:
   - Go to Azure Portal > Azure Active Directory
   - Navigate to App registrations
   - Create a new registration
   - Add Microsoft Graph API permissions for Mail.ReadWrite

2. Install the package:
```bash
npm install exchange365-draft-emails
```

3. Create a .env file with your credentials:
```env
TENANT_ID=your_tenant_id
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
USER_EMAIL=your_email@domain.com
```

## Usage

```javascript
const Exchange365DraftEmails = require('exchange365-draft-emails');

const config = {
    tenantId: process.env.TENANT_ID,
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    userEmail: process.env.USER_EMAIL
};

const emailClient = new Exchange365DraftEmails(config);

// Create a draft email
const draftEmail = {
    subject: 'Test Draft Email',
    body: 'This is a test draft email',
    to: ['recipient@example.com']
};

emailClient.createDraft(draftEmail)
    .then(result => console.log(result))
    .catch(error => console.error(error));

// Get all draft emails
emailClient.getAllDrafts()
    .then(result => console.log(result))
    .catch(error => console.error(error));
```

## Features

- Create draft emails in Exchange/Office 365
- Retrieve all draft emails
- Uses Microsoft Graph API for secure communication
- Support for text and HTML content
- Error handling and logging

## Authentication

This library uses client credentials flow with Azure AD. Make sure your application has the necessary permissions in Azure AD:
- Mail.ReadWrite
- Mail.Send (if you plan to add send functionality)