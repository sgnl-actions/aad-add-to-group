/**
 * Azure Active Directory Add User to Group Action
 *
 * Adds a user to a group in Azure Active Directory using Microsoft Graph API.
 */

import { getBaseURL, createAuthHeaders } from '@sgnl-actions/utils';

/**
 * Helper function to check if user exists in Azure AD
 * @param {string} userPrincipalName - User Principal Name (UPN) of the user
 * @param {string} baseUrl - Azure AD base URL
 * @param {Object} headers - Request headers with Authorization
 * @returns {Promise<boolean>} - True if user exists, false otherwise
 */
async function userExists(userPrincipalName, baseUrl, headers) {
  const encodedUPN = encodeURIComponent(userPrincipalName);
  const url = `${baseUrl}/v1.0/users/${encodedUPN}`;

  const response = await fetch(url, {
    method: 'GET',
    headers
  });

  return response.ok;
}

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

    console.log(`Checking if user ${userPrincipalName} exists in Azure AD`);

    try {
      // First, check if user exists in Azure AD
      const doesUserExist = await userExists(
        userPrincipalName,
        baseUrl,
        headers
      );

      if (!doesUserExist) {
        throw new Error(`User ${userPrincipalName} does not exist in Azure AD`);
      }

      console.log(`User ${userPrincipalName} exists. Adding to group ${groupId}`);

      // Try to add the user to the group
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
          message: 'User successfully added to group',
          address: baseUrl
        };
      } else if (response.status === 400) {
        // Check if this is a "already exists" error
        const errorData = await response.json().catch(() => null);
        if (errorData && errorData.error && errorData.error.code === 'Request_BadRequest' &&
            errorData.error.message && errorData.error.message.includes('already exist')) {
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

        // If it's a different 400 error, throw it
        const errorText = await response.text().catch(() => 'Unknown error');
        throw new Error(`Failed to add user to group: ${response.status} ${response.statusText} - ${errorText}`);
      } else {
        const errorText = await response.text().catch(() => 'Unknown error');
        throw new Error(`Failed to add user to group: ${response.status} ${response.statusText} - ${errorText}`);
      }
    } catch (error) {
      console.error(`Error in group membership operation: ${error.message}`);
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