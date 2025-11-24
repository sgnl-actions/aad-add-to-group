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

### Required Secrets

- **`BEARER_AUTH_TOKEN`**: Bearer token for Azure AD API authentication

### Required Environment Variables

- **`AZURE_AD_TENANT_URL`**: Azure AD tenant URL (e.g., `https://graph.microsoft.com`)

### Input Parameters

- **`userPrincipalName`** (required): User Principal Name (UPN) of the user to add to the group (e.g., `user@domain.com`)
- **`groupId`** (required): Azure AD Group ID (GUID format, e.g., `12345678-1234-1234-1234-123456789012`)

### Output Parameters

- **`status`**: Operation result (`success`, `failed`, `recovered`, etc.)
- **`userPrincipalName`**: User Principal Name that was processed
- **`groupId`**: Azure AD Group ID that was processed
- **`added`**: Boolean indicating whether the user was successfully added to the group
- **`message`**: Optional message providing additional context (e.g., when user is already a member)

## Usage Examples

### Basic Usage

```json
{
  "userPrincipalName": "john.doe@company.com",
  "groupId": "12345678-1234-1234-1234-123456789012"
}
```

### In a SGNL Job Specification

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
    "AZURE_AD_TENANT_URL": "https://graph.microsoft.com"
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

- **Authentication**: Uses Bearer token authentication with Azure AD
- **URL Encoding**: All user principal names and group IDs are properly URL encoded to prevent injection attacks
- **Input Validation**: Validates required parameters and environment variables
- **Token Security**: Authentication tokens are never logged or exposed

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

3. **"BEARER_AUTH_TOKEN secret is required"**
   - Ensure the authentication token is configured in secrets
   - Verify the token has not expired

4. **"Authentication failed: 401"**
   - Check if the Azure AD token is valid and not expired
   - Verify the application has proper permissions in Azure AD

5. **"Authentication failed: 403"**
   - Verify the application has the required Microsoft Graph permissions:
     - `GroupMember.ReadWrite.All`
     - `User.Read.All`
     - `Group.Read.All`

6. **"Bad request: Invalid user ID"**
   - Verify the user exists in Azure AD
   - Check that the UPN format is correct

7. **"Bad request: Invalid group ID"**
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