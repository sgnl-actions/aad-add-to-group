/**
 * Azure Active Directory Add User to Group Action
 *
 * Adds a user to a group in Azure Active Directory using Microsoft Graph API.
 */

import { getBaseURL, createAuthHeaders, resolveJSONPathTemplates} from '@sgnl-actions/utils';

/**
 * Helper function to add a user to a group in Azure AD
 * @param {string} userPrincipalName - User Principal Name (UPN) of the user
 * @param {string} groupId - Azure AD Group ID (GUID)
 * @param {string} baseUrl - Azure AD base URL
 * @param {Object} headers - Request headers with Authorization
 * @returns {Promise<Response>} - Fetch response object
 */
async function addUserToGroup(userPrincipalName, groupId, baseUrl, headers) {
  // URL encode the user principal name for the OData reference
  const encodedUPN = encodeURIComponent(userPrincipalName);

  // URL encode the group ID to handle any special characters
  const encodedGroupId = encodeURIComponent(groupId);

  // Construct the Graph API endpoint
  const url = `${baseUrl}/v1.0/groups/${encodedGroupId}/members/$ref`;

  // Prepare the request body with OData reference to the user
  const requestBody = {
    '@odata.id': `https://graph.microsoft.com/v1.0/users/${encodedUPN}`
  };

  const response = await fetch(url, {
    method: 'POST',
    headers,
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
   *
   * @returns {Object} Job results
   */
  invoke: async (params, context) => {
    console.log('Starting Azure AD add user to group operation');

    const jobContext = context.data || {};

    // Resolve JSONPath templates in params
    const { result: resolvedParams, errors } = resolveJSONPathTemplates(params, jobContext);
    if (errors.length > 0) {
      console.warn('Template resolution errors:', errors);
    }

    const { userPrincipalName, groupId } = resolvedParams;

    // Get base URL and authentication headers using utilities
    const baseUrl = getBaseURL(resolvedParams, context);
    const headers = await createAuthHeaders(context);

    console.log(`Adding user ${userPrincipalName} to group ${groupId}`);

    try {
      const response = await addUserToGroup(
        userPrincipalName,
        groupId,
        baseUrl,
        headers
      );

      // Handle different response scenarios
      if (response.status === 204) {
        // Success - user added to group
        console.log(`Successfully added user ${userPrincipalName} to group ${groupId}`);
        return {
          status: 'success',
          userPrincipalName,
          groupId,
          added: true,
          address: baseUrl
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
            message: 'User is already a member of the group',
            address: baseUrl
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
   * Error recovery handler - framework handles retries by default
   * Only implement if custom recovery logic is needed
   * @param {Object} params - Original params plus error information
   * @param {Object} context - Execution context
   * @returns {Object} Recovery results
   */
  error: async (params, _context) => {
    const { error, userPrincipalName, groupId } = params;
    console.error(`User group assignment failed for user ${userPrincipalName} to group ${groupId}: ${error.message}`);

    // Framework handles retries for transient errors (429, 502, 503, 504)
    // Just re-throw the error to let the framework handle it
    throw error;
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