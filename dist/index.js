// SGNL Job Script - Auto-generated bundle
'use strict';

/**
 * Azure Active Directory Add User to Group Action
 *
 * Adds a user to a group in Azure Active Directory using Microsoft Graph API.
 */

/**
 * Helper function to add a user to a group in Azure AD
 * @param {string} userPrincipalName - User Principal Name (UPN) of the user
 * @param {string} groupId - Azure AD Group ID (GUID)
 * @param {string} tenantUrl - Azure AD tenant URL
 * @param {string} authToken - Azure AD authentication token
 * @returns {Promise<Response>} - Fetch response object
 */
async function addUserToGroup(userPrincipalName, groupId, tenantUrl, authToken) {
  // URL encode the user principal name for the OData reference
  const encodedUPN = encodeURIComponent(userPrincipalName);

  // URL encode the group ID to handle any special characters
  const encodedGroupId = encodeURIComponent(groupId);

  // Construct the Graph API endpoint
  const url = new URL(`/v1.0/groups/${encodedGroupId}/members/$ref`, tenantUrl);

  // Prepare the request body with OData reference to the user
  const requestBody = {
    '@odata.id': `https://graph.microsoft.com/v1.0/users/${encodedUPN}`
  };

  const response = await fetch(url.toString(), {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${authToken}`,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    body: JSON.stringify(requestBody)
  });

  return response;
}

var script = {
  /**
   * Main execution handler - adds a user to a group in Azure AD
   * @param {Object} params - Job input parameters
   * @param {string} params.userPrincipalName - User Principal Name (UPN)
   * @param {string} params.groupId - Azure AD Group ID (GUID)
   * @param {Object} context - Execution context with env, secrets, outputs
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

    // Validate required environment variables
    if (!context.environment?.AZURE_AD_TENANT_URL) {
      throw new Error('AZURE_AD_TENANT_URL environment variable is required');
    }

    // Validate required secrets
    if (!context.secrets?.AZURE_AD_TOKEN) {
      throw new Error('AZURE_AD_TOKEN secret is required');
    }

    console.log(`Adding user ${userPrincipalName} to group ${groupId}`);

    try {
      const response = await addUserToGroup(
        userPrincipalName,
        groupId,
        context.environment.AZURE_AD_TENANT_URL,
        context.secrets.AZURE_AD_TOKEN
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
        const response = await addUserToGroup(
          params.userPrincipalName,
          params.groupId,
          context.environment.AZURE_AD_TENANT_URL,
          context.secrets.AZURE_AD_TOKEN
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

module.exports = script;
