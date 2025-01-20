const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
require('dotenv').config();

class Exchange365DraftEmails {
    constructor(config) {
        this.tenantId = config.tenantId;
        this.clientId = config.clientId;
        this.clientSecret = config.clientSecret;
        this.userEmail = config.userEmail;
        
        this.credential = new ClientSecretCredential(
            this.tenantId,
            this.clientId,
            this.clientSecret
        );
        
        this.client = Client.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => {
                    const token = await this.credential.getToken('https://graph.microsoft.com/.default');
                    return token.token;
                }
            }
        });
    }

    async createDraft(emailDetails) {
        try {
            const message = {
                subject: emailDetails.subject,
                body: {
                    contentType: 'text',
                    content: emailDetails.body
                },
                toRecipients: emailDetails.to.map(recipient => ({
                    emailAddress: {
                        address: recipient
                    }
                })),
                isDraft: true
            };

            await this.client.api(`/users/${this.userEmail}/messages`)
                .post(message);

            return {
                success: true,
                message: 'Draft email created successfully'
            };
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    async getAllDrafts() {
        try {
            const drafts = await this.client.api(`/users/${this.userEmail}/messages`)
                .filter("isDraft eq true")
                .get();

            return {
                success: true,
                drafts: drafts.value
            };
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }
}

module.exports = Exchange365DraftEmails;