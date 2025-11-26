/**
 * Azure Active Directory Add User to Group Action
 *
 * Adds a user to a group in Azure Active Directory using Microsoft Graph API.
 */

/**
 * Get OAuth2 access token using client credentials flow
 * @param {Object} config - OAuth2 configuration
 * @returns {Promise<string>} Access token
 */
async function getClientCredentialsToken(config) {
  const { tokenUrl, clientId, clientSecret, scope, audience, authStyle } = config;

  const params = new URLSearchParams();
  params.append('grant_type', 'client_credentials');

  if (scope) {
    params.append('scope', scope);
  }

  if (audience) {
    params.append('audience', audience);
  }

  const headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Accept': 'application/json'
  };

  if (authStyle === 'InParams') {
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
  } else {
    // InHeader, AutoDetect, or undefined
    const credentials = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
    headers['Authorization'] = `Basic ${credentials}`;
  }

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers,
    body: params.toString()
  });

  if (!response.ok) {
    let errorText;
    try {
      const errorData = await response.json();
      errorText = JSON.stringify(errorData);
    } catch {
      errorText = await response.text();
    }
    throw new Error(
      `OAuth2 token request failed: ${response.status} ${response.statusText} - ${errorText}`
    );
  }

  const data = await response.json();
  
  if (!data.access_token) {
    throw new Error('No access_token in OAuth2 response');
  }

  return data.access_token;
}

/**
 * Helper function to add a user to a group in Azure AD
 * @param {string} userPrincipalName - User Principal Name (UPN) of the user
 * @param {string} groupId - Azure AD Group ID (GUID)
 * @param {string} address - Azure AD base URL
 * @param {string} accessToken - OAuth2 access token
 * @returns {Promise<Response>} - Fetch response object
 */
async function addUserToGroup(userPrincipalName, groupId, address, accessToken) {
  // Remove trailing slash from address if present
  const cleanAddress = address.endsWith('/') ? address.slice(0, -1) : address;

  // URL encode the user principal name for the OData reference
  const encodedUPN = encodeURIComponent(userPrincipalName);

  // URL encode the group ID to handle any special characters
  const encodedGroupId = encodeURIComponent(groupId);

  // Construct the Graph API endpoint
  const url = `${cleanAddress}/v1.0/groups/${encodedGroupId}/members/$ref`;

  // Prepare the request body with OData reference to the user
  const requestBody = {
    '@odata.id': `https://graph.microsoft.com/v1.0/users/${encodedUPN}`
  };

  const authHeader = accessToken.startsWith('Bearer ') ? accessToken : `Bearer ${accessToken}`;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': authHeader,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    body: JSON.stringify(requestBody)
  });

  return response;
}

export default {
  /**
   * Main execution handler - adds a user to a group in Azure AD
   * @param {Object} params - Job input parameters
   * @param {string} params.userPrincipalName - User Principal Name (UPN)
   * @param {string} params.groupId - Azure AD Group ID (GUID)
   * @param {string} params.address - The Azure AD API base URL (e.g., https://graph.microsoft.com)
   * @param {Object} context - Execution context with env, secrets, outputs
   * @param {string} context.environment.ADDRESS - Default Azure AD API base URL
   *
   * The configured auth type will determine which of the following environment variables and secrets are available
   * @param {string} context.secrets.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_AUDIENCE
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_AUTH_STYLE
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_SCOPE
   * @param {string} context.environment.OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL
   *
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_AUTHORIZATION_CODE
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_CLIENT_SECRET
   * @param {string} context.secrets.OAUTH2_AUTHORIZATION_CODE_REFRESH_TOKEN
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_AUTH_STYLE
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_AUTH_URL
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_CLIENT_ID
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_LAST_TOKEN_ROTATION_TIMESTAMP
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_REDIRECT_URI
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_SCOPE
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_LIFETIME_FREQUENCY
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_ROTATION_FREQUENCY
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_ROTATION_INTERVAL
   * @param {string} context.environment.OAUTH2_AUTHORIZATION_CODE_TOKEN_URL
   *
   * @returns {Object} Job results
   */
  invoke: async (params, context) => {
    console.log('Starting Azure AD add user to group operation');

    // Validate required inputs
    const { userPrincipalName, groupId } = params;

    if (!userPrincipalName) {
      throw new Error('userPrincipalName is required');
    }

    if (!groupId) {
      throw new Error('groupId is required');
    }

    // Determine the URL to use
    const address = params.address || context.environment?.ADDRESS;
    if (!address) {
      throw new Error('No URL specified. Provide either address parameter or ADDRESS environment variable');
    }

    let accessToken;

    if (context.secrets?.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN) {
      accessToken = context.secrets.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN;
    } else if (context.secrets?.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET) {
      const tokenUrl = context.environment?.OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL;
      const clientId = context.environment?.OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID;
      const clientSecret = context.secrets.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET;

      if (!tokenUrl || !clientId || !clientSecret) {
        throw new Error('OAuth2 Client Credentials flow requires TOKEN_URL, CLIENT_ID, and CLIENT_SECRET');
      }

      accessToken = await getClientCredentialsToken({
        tokenUrl,
        clientId,
        clientSecret,
        scope: context.environment?.OAUTH2_CLIENT_CREDENTIALS_SCOPE,
        audience: context.environment?.OAUTH2_CLIENT_CREDENTIALS_AUDIENCE,
        authStyle: context.environment?.OAUTH2_CLIENT_CREDENTIALS_AUTH_STYLE
      });
    } else {
      throw new Error('OAuth2 authentication is required. Configure either Authorization Code or Client Credentials flow.');
    }

    console.log(`Adding user ${userPrincipalName} to group ${groupId}`);

    try {
      const response = await addUserToGroup(
        userPrincipalName,
        groupId,
        address,
        accessToken
      );

      // Handle different response scenarios
      if (response.status === 204) {
        // Success - user added to group
        console.log(`Successfully added user ${userPrincipalName} to group ${groupId}`);
        return {
          status: 'success',
          userPrincipalName,
          groupId,
          added: true
        };
      } else if (response.status === 400) {
        // Bad request - could be user already in group or invalid IDs
        const errorText = await response.text();
        if (errorText.includes('already a member')) {
          console.log(`User ${userPrincipalName} is already a member of group ${groupId}`);
          return {
            status: 'success',
            userPrincipalName,
            groupId,
            added: false,
            message: 'User is already a member of the group'
          };
        }
        throw new Error(`Bad request: ${errorText}`);
      } else {
        // Other error responses
        const errorText = await response.text();
        throw new Error(`Failed to add user to group: ${response.status} ${response.statusText} - ${errorText}`);
      }
    } catch (error) {
      console.error(`Error adding user to group: ${error.message}`);
      throw error;
    }
  },

  /**
   * Error recovery handler - handles retry logic for transient failures
   * @param {Object} params - Original params plus error information
   * @param {Object} context - Execution context
   * @returns {Object} Recovery results
   */
  error: async (params, context) => {
    const { error } = params;
    console.error(`Azure AD add user to group encountered error: ${error.message}`);

    // Check for retryable errors (rate limiting, server errors)
    if (error.message.includes('429') ||
        error.message.includes('502') ||
        error.message.includes('503') ||
        error.message.includes('504')) {

      console.log('Retryable error detected, waiting before retry...');
      await new Promise(resolve => setTimeout(resolve, 5000)); // 5 second delay

      // Attempt recovery by retrying the operation
      try {
        const address = params.address || context.environment?.ADDRESS;

        let accessToken;
        if (context.secrets?.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN) {
          accessToken = context.secrets.OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN;
        } else if (context.secrets?.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET) {
          accessToken = await getClientCredentialsToken({
            tokenUrl: context.environment?.OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL,
            clientId: context.environment?.OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID,
            clientSecret: context.secrets.OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET,
            scope: context.environment?.OAUTH2_CLIENT_CREDENTIALS_SCOPE,
            audience: context.environment?.OAUTH2_CLIENT_CREDENTIALS_AUDIENCE,
            authStyle: context.environment?.OAUTH2_CLIENT_CREDENTIALS_AUTH_STYLE
          });
        }

        const response = await addUserToGroup(
          params.userPrincipalName,
          params.groupId,
          address,
          accessToken
        );

        if (response.status === 204) {
          console.log(`Recovery successful: user ${params.userPrincipalName} added to group ${params.groupId}`);
          return {
            status: 'recovered',
            userPrincipalName: params.userPrincipalName,
            groupId: params.groupId,
            added: true
          };
        }
      } catch (recoveryError) {
        console.error(`Recovery attempt failed: ${recoveryError.message}`);
      }
    }

    // For authentication errors (401, 403) or other permanent failures, don't retry
    if (error.message.includes('401') || error.message.includes('403')) {
      console.error('Authentication error - operation cannot be retried');
      throw error;
    }

    // Default: let framework handle retry
    return { status: 'retry_requested' };
  },

  /**
   * Graceful shutdown handler - performs cleanup
   * @param {Object} params - Original params plus halt reason
   * @param {Object} context - Execution context
   * @returns {Object} Cleanup results
   */
  halt: async (params, _context) => {
    const { reason, userPrincipalName, groupId } = params;
    console.log(`Azure AD add user to group operation halted: ${reason}`);

    // No specific cleanup needed for this operation
    return {
      status: 'halted',
      userPrincipalName: userPrincipalName || 'unknown',
      groupId: groupId || 'unknown',
      reason: reason,
      halted_at: new Date().toISOString()
    };
  }
};