# Azure AD Add User to Group Action

This action adds a user to a group in Azure Active Directory using the Microsoft Graph API.

## Overview

The Azure AD Add User to Group action enables automated group membership management by adding users to Azure AD security groups or Microsoft 365 groups. It handles URL encoding, error scenarios, and provides comprehensive retry logic for reliable execution.

## Prerequisites

- Azure AD tenant with appropriate permissions
- Application registered in Azure AD with the following Microsoft Graph permissions:
  - `GroupMember.ReadWrite.All` (to add users to groups)
  - `User.Read.All` (to validate user existence)
  - `Group.Read.All` (to validate group existence)

## Configuration

### Authentication

This action supports two OAuth2 authentication methods:

#### Option 1: OAuth2 Client Credentials
| Secret/Environment | Required | Description |
|-------------------|----------|-------------|
| `OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET` | Yes | OAuth2 client secret |
| `OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID` | Yes | OAuth2 client ID |
| `OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL` | Yes | OAuth2 token endpoint URL |
| `OAUTH2_CLIENT_CREDENTIALS_SCOPE` | No | OAuth2 scope |
| `OAUTH2_CLIENT_CREDENTIALS_AUDIENCE` | No | OAuth2 audience |
| `OAUTH2_CLIENT_CREDENTIALS_AUTH_STYLE` | No | OAuth2 auth style |

#### Option 2: OAuth2 Authorization Code
| Secret | Required | Description |
|--------|----------|-------------|
| `OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN` | Yes | OAuth2 access token |

### Environment Variables

| Variable | Required | Description | Example |
|----------|----------|-------------|---------|
| `ADDRESS` | Yes | Default Azure AD API base URL | `https://graph.microsoft.com` |

### Input Parameters

| Parameter | Type | Required | Description | Example |
|-----------|------|----------|-------------|---------|
| `userPrincipalName` | string | Yes | User Principal Name (UPN) of the user to add | `user@domain.com` |
| `groupId` | string | Yes | Azure AD Group ID (GUID) | `12345678-1234-1234-1234-123456789012` |
| `address` | string | No | Optional Azure AD API base URL override | `https://graph.microsoft.com` |

### Output Structure

| Field | Type | Description |
|-------|------|-------------|
| `status` | string | Operation result (success, failed, etc.) |
| `userPrincipalName` | string | User Principal Name that was processed |
| `groupId` | string | Azure AD Group ID that was processed |
| `added` | boolean | Whether the user was successfully added to the group |
| `address` | string | The Azure AD API base URL used |
| `message` | string | Optional message providing additional context (e.g., when user is already a member)

## Usage Examples

### Basic Usage

```json
{
  "userPrincipalName": "john.doe@company.com",
  "groupId": "12345678-1234-1234-1234-123456789012"
}
```

### With OAuth2 Client Credentials

```json
{
  "id": "add-user-to-hr-group",
  "type": "nodejs-22",
  "script": {
    "repository": "github.com/sgnl-actions/aad-add-to-group",
    "version": "v1.0.0",
    "type": "nodejs"
  },
  "script_inputs": {
    "userPrincipalName": "new.employee@company.com",
    "groupId": "a1b2c3d4-e5f6-7890-1234-56789abcdef0"
  },
  "environment": {
    "ADDRESS": "https://graph.microsoft.com",
    "OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL": "https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token",
    "OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID": "your-client-id",
    "OAUTH2_CLIENT_CREDENTIALS_SCOPE": "https://graph.microsoft.com/.default"
  },
  "secrets": {
    "OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET": "your-client-secret"
  }
}
```

### With OAuth2 Authorization Code

```json
{
  "id": "add-user-to-hr-group",
  "type": "nodejs-22",
  "script": {
    "repository": "github.com/sgnl-actions/aad-add-to-group",
    "version": "v1.0.0",
    "type": "nodejs"
  },
  "script_inputs": {
    "userPrincipalName": "new.employee@company.com",
    "groupId": "a1b2c3d4-e5f6-7890-1234-56789abcdef0"
  },
  "environment": {
    "ADDRESS": "https://graph.microsoft.com"
  },
  "secrets": {
    "OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN": "your-access-token"
  }
}
```

## API Details

This action uses the Microsoft Graph API endpoint:

```
POST https://graph.microsoft.com/v1.0/groups/{groupId}/members/$ref
```

The request body uses the OData reference format:

```json
{
  "@odata.id": "https://graph.microsoft.com/v1.0/users/{encodedUserPrincipalName}"
}
```

## Error Handling

### Success Scenarios

- **204 No Content**: User successfully added to group
- **400 Bad Request** (user already member): Treated as success with `added: false`

### Retryable Errors

The action automatically retries on:
- **429 Too Many Requests**: Rate limiting
- **502 Bad Gateway**: Server error
- **503 Service Unavailable**: Temporary service issues
- **504 Gateway Timeout**: Request timeout

### Fatal Errors

The following errors will not be retried:
- **401 Unauthorized**: Invalid or expired authentication token
- **403 Forbidden**: Insufficient permissions
- **400 Bad Request** (other than "already member"): Invalid user ID or group ID

### Recovery Logic

The error handler implements exponential backoff with a 5-second initial delay for retryable errors. It attempts one recovery operation before falling back to the framework's retry mechanism.

## Security Considerations

- **Authentication**: Uses OAuth2 authentication (Authorization Code or Client Credentials flow)
- **Token Management**: Access tokens are automatically fetched for Client Credentials flow
- **URL Encoding**: All user principal names and group IDs are properly URL encoded to prevent injection attacks
- **Input Validation**: Validates required parameters and environment variables
- **Token Security**: Authentication tokens and secrets are never logged or exposed

## Development

### Local Testing

```bash
# Run with mock parameters
npm run dev -- --params '{"userPrincipalName": "test@example.com", "groupId": "12345678-1234-1234-1234-123456789012"}'

# Run unit tests
npm test

# Check test coverage
npm run test:coverage
```

### Building

```bash
# Build distribution bundle
npm run build

# Lint code
npm run lint
```

## Troubleshooting

### Common Issues

1. **"userPrincipalName is required"**
   - Ensure the `userPrincipalName` parameter is provided
   - Verify the UPN format matches your Azure AD configuration

2. **"groupId is required"**
   - Ensure the `groupId` parameter is provided
   - Verify the group ID is a valid GUID format

3. **"OAuth2 authentication is required"**
   - Ensure either OAuth2 Authorization Code or Client Credentials flow is configured
   - Verify all required secrets and environment variables are set

4. **"OAuth2 Client Credentials flow requires TOKEN_URL, CLIENT_ID, and CLIENT_SECRET"**
   - Check that `OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL` is set
   - Check that `OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID` is set
   - Check that `OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET` is set in secrets

5. **"OAuth2 token request failed"**
   - Verify the token URL is correct
   - Check that client credentials are valid
   - Ensure the scope is appropriate for Microsoft Graph API

6. **"Authentication failed: 401"**
   - Check if the OAuth2 access token is valid and not expired
   - Verify the application has proper permissions in Azure AD

7. **"Authentication failed: 403"**
   - Verify the application has the required Microsoft Graph permissions:
     - `GroupMember.ReadWrite.All`
     - `User.Read.All`
     - `Group.Read.All`

8. **"Bad request: Invalid user ID"**
   - Verify the user exists in Azure AD
   - Check that the UPN format is correct

9. **"Bad request: Invalid group ID"**
   - Verify the group exists in Azure AD
   - Check that the group ID is a valid GUID

### Rate Limiting

Microsoft Graph API has rate limits. The action includes automatic retry logic with exponential backoff for rate limit responses (429). If you encounter persistent rate limiting:

1. Consider implementing delays between batch operations
2. Check if other processes are making concurrent API calls
3. Review your Azure AD application's API usage patterns

### Testing Group Membership

To verify the action worked correctly, you can check group membership using:

```bash
# Using Microsoft Graph API
curl -H "Authorization: Bearer $TOKEN" \
  "https://graph.microsoft.com/v1.0/groups/$GROUP_ID/members"

# Using Azure CLI
az ad group member list --group $GROUP_ID --query "[?userPrincipalName=='$USER_UPN']"
```

## Support

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Azure AD Group Management](https://docs.microsoft.com/en-us/graph/api/group-post-members)
- [SGNL Actions Documentation](https://github.com/sgnl-actions)