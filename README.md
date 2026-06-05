# communication-backend

Unified standalone backend service for Alfred communication workflows:

- Send email (text/html + PNG/PDF attachments)
- Slack OAuth token persistence in AWS SSM (with parameter reference in DynamoDB)
- Slack messaging (text + PNG/PDF attachments)
- Teams OAuth token persistence in AWS SSM (with parameter reference in DynamoDB) and messaging (text + optional file metadata)

## APIs

- GET /health

### Email
- POST /api/v1/email/send

Request body:

```json
{
  "to": ["user@example.com"],
  "subject": "Quarterly report",
  "text": "Please review the attached PDF.",
  "html": "<p>Please review the attached PDF.</p>",
  "attachments": [
    {
      "fileName": "report.pdf",
      "mimeType": "application/pdf",
      "contentBase64": "BASE64_CONTENT"
    }
  ]
}
```

Attachment mime types supported:
- image/png
- application/pdf

### Slack OAuth
- GET /api/v1/slack/oauth/install-url?clientContext=optional
- GET /api/v1/slack/oauth/callback
- GET /api/v1/slack/oauth/status?organizationId=...

`organizationId` is required on install-url so callback stores tokens for the right tenant:

```json
{
  "organizationId": "org_123"
}
```

### Slack Messaging
- GET /api/v1/slack/channels
- GET /api/v1/slack/dms
- POST /api/v1/slack/send

For GET endpoints, provide bot token in Authorization header:
- query param `organizationId=<orgId>` or header `x-organization-id: <orgId>`

Send request body:

```json
{
  "channelId": "C0123456789",
  "message": "Hello from Alfred",
  "file": {
    "fileName": "chart.png",
    "mimeType": "image/png",
    "contentBase64": "BASE64_CONTENT"
  }
}
```

`message` and `file` are optional individually, but at least one is required.

### Teams OAuth
- GET /api/v1/teams/oauth/auth-url?clientContext=optional
- GET /api/v1/teams/oauth/callback
- GET /api/v1/teams/oauth/status?organizationId=...

Teams callback redirects to the frontend origin with non-sensitive query parameters:
- teamsStatus=success|error
- teamsExpiresAt

### Teams Messaging
- GET /api/v1/teams/chats
- GET /api/v1/teams/channels
- POST /api/v1/teams/send

For GET endpoints, provide Teams token in Authorization header:
- query param `organizationId=<orgId>` or header `x-organization-id: <orgId>`

Send request body:

```json
{
  "destinationType": "chat",
  "destinationId": "19:...",
  "teamId": "optional-for-channel",
  "message": "Hello from Alfred",
  "file": {
    "fileName": "report.pdf",
    "mimeType": "application/pdf",
    "contentBase64": "BASE64_CONTENT"
  }
}
```

`message` and `file` are optional individually, but at least one is required.

## Environment Variables

Env variables merged from existing Alfred services:

- PORT
- FRONTEND_ORIGIN
- FRONTEND_CALLBACK_URL

From alfred-backend email setup:
- SENDGRID_API_KEY
- FROM_EMAIL
- FROM_NAME

From existing Slack integration/slack-token-backend:
- SLACK_CLIENT_ID
- SLACK_CLIENT_SECRET
- SLACK_REDIRECT_URI
- SLACK_SCOPES
- SLACK_USER_SCOPES
- SLACK_BASE_URL
- STATE_SECRET
- TOKEN_TTL_MS

For server-side token storage:
- AWS_REGION
- COMMUNICATION_TOKENS_TABLE_NAME
- COMMUNICATION_TOKENS_SSM_PREFIX
- COMMUNICATION_TOKENS_KMS_KEY_ID (optional)

## Credential Storage Design

- DynamoDB table key schema:
  - partition key: `organizationId` (String)
  - sort key: `provider` (String, values: `slack`, `teams`)
- Stored item fields:
  - `parameterName` (SSM SecureString path)
  - `updatedAt`
  - `expiresAt` (optional)
  - `metadata` (provider metadata, non-secret)
- SSM SecureString value:
  - JSON payload containing provider tokens
  - Slack: bot/user token bundle
  - Teams: access token + expiry

For Teams integration:
- TEAMS_CLIENT_ID
- TEAMS_CLIENT_SECRET
- TEAMS_REDIRECT_URI
- TEAMS_SCOPES (comma-separated)

## Run

```bash
npm install
npm run dev
```

## Build + Start

```bash
npm run build
npm start
```
