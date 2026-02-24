/**
 * Azure Active Directory Add User to Group Action
 *
 * Adds a user to a group in Azure Active Directory using Microsoft Graph API.
 */

import { getBaseURL, createAuthHeaders } from '@sgnl-actions/utils';

/**
 * Helper function to add a user to a group in Azure AD
 * @param {string} userPrincipalName - User Principal Name (UPN) of the user
 * @param {string} groupId - Azure AD Group ID (GUID)
 * @param {string} baseUrl - Azure AD base URL
 * @param {Object} headers - Request headers with Authorization
 * @returns {Promise<Response>} - Fetch response object
 */
async function addUserToGroup(userPrincipalName, groupId, baseUrl, headers) {
  const encodedUPN = encodeURIComponent(userPrincipalName);
  const encodedGroupId = encodeURIComponent(groupId);
  const url = `${baseUrl}/v1.0/groups/${encodedGroupId}/members/$ref`;

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

    const { userPrincipalName, groupId } = params;

    // Validate required inputs before making any API calls
    if (!userPrincipalName || typeof userPrincipalName !== 'string' || !userPrincipalName.trim()) {
      throw new Error('userPrincipalName parameter is required and cannot be empty');
    }
    if (!groupId || typeof groupId !== 'string' || !groupId.trim()) {
      throw new Error('groupId parameter is required and cannot be empty');
    }

    // Get base URL and authentication headers using utilities
    const baseUrl = getBaseURL(params, context);
    const headers = await createAuthHeaders(context);

    console.log(`Adding user ${userPrincipalName} to group ${groupId}`);

    try {
      const response = await addUserToGroup(
        userPrincipalName,
        groupId,
        baseUrl,
        headers
      );

      if (response.status === 204) {
        console.log(`Successfully added user ${userPrincipalName} to group ${groupId}`);
        return {
          status: 'success',
          userPrincipalName,
          groupId,
          added: true,
          address: baseUrl
        };
      } else if (response.status === 400) {
        const errorText = await response.text();
        // Handle both possible Azure "already a member" error message formats
        if (errorText.includes('already a member') || errorText.includes("modified properties: 'members'")) {
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
        const errorText = await response.text();
        throw new Error(`Failed to add user to group: ${response.status} ${response.statusText} - ${errorText}`);
      }
    } catch (error) {
      console.error(`Error adding user to group: ${error.message}`);
      throw error;
    }
  },

  error: async (params, _context) => {
    const { error, userPrincipalName, groupId } = params;
    console.error(`User group assignment failed for user ${userPrincipalName} to group ${groupId}: ${error.message}`);
    throw error;
  },

  halt: async (params, _context) => {
    const { reason, userPrincipalName, groupId } = params;
    console.log(`Azure AD add user to group operation halted: ${reason}`);

    return {
      status: 'halted',
      userPrincipalName: userPrincipalName || 'unknown',
      groupId: groupId || 'unknown',
      reason: reason,
      halted_at: new Date().toISOString()
    };
  }
};