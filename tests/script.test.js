import { jest } from '@jest/globals';
import script from '../src/script.mjs';

// Mock fetch globally
global.fetch = jest.fn();
global.URL = URL;

describe('Azure AD Add User to Group Script', () => {
  const mockContext = {
    environment: {
      ADDRESS: 'https://graph.microsoft.com'
    },
    secrets: {
      OAUTH2_AUTHORIZATION_CODE_ACCESS_TOKEN: 'test-token-123456'
    }
  };

  beforeEach(() => {
    jest.clearAllMocks();
    global.console.log = jest.fn();
    global.console.error = jest.fn();
  });

  describe('invoke handler', () => {
    test('should successfully add user to group (204 response)', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      global.fetch.mockResolvedValueOnce({
        status: 204,
        ok: true
      });

      const result = await script.invoke(params, mockContext);

      expect(result.status).toBe('success');
      expect(result.userPrincipalName).toBe('test-user@example.com');
      expect(result.groupId).toBe('12345678-1234-1234-1234-123456789012');
      expect(result.added).toBe(true);

      // Verify fetch was called correctly
      expect(global.fetch).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/groups/12345678-1234-1234-1234-123456789012/members/$ref',
        {
          method: 'POST',
          headers: {
            'Authorization': 'Bearer test-token-123456',
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify({
            '@odata.id': 'https://graph.microsoft.com/v1.0/users/test-user%40example.com'
          })
        }
      );
    });

    test('should handle user already in group (400 response)', async () => {
      const params = {
        userPrincipalName: 'existing-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      global.fetch.mockResolvedValueOnce({
        status: 400,
        ok: false,
        text: jest.fn().mockResolvedValue('User is already a member of the group')
      });

      const result = await script.invoke(params, mockContext);

      expect(result.status).toBe('success');
      expect(result.userPrincipalName).toBe('existing-user@example.com');
      expect(result.groupId).toBe('12345678-1234-1234-1234-123456789012');
      expect(result.added).toBe(false);
      expect(result.message).toBe('User is already a member of the group');
    });

    test('should URL encode user principal name with special characters', async () => {
      const params = {
        userPrincipalName: 'test+user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      global.fetch.mockResolvedValueOnce({
        status: 204,
        ok: true
      });

      await script.invoke(params, mockContext);

      // Verify the encoded UPN in the OData reference
      expect(global.fetch).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          body: JSON.stringify({
            '@odata.id': 'https://graph.microsoft.com/v1.0/users/test%2Buser%40example.com'
          })
        })
      );
    });

    test('should URL encode group ID with special characters', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: 'group id with spaces'
      };

      global.fetch.mockResolvedValueOnce({
        status: 204,
        ok: true
      });

      await script.invoke(params, mockContext);

      // Verify the encoded group ID in the URL
      expect(global.fetch).toHaveBeenCalledWith(
        'https://graph.microsoft.com/v1.0/groups/group%20id%20with%20spaces/members/$ref',
        expect.any(Object)
      );
    });

    test('should throw error for missing userPrincipalName', async () => {
      const params = {
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      await expect(script.invoke(params, mockContext)).rejects.toThrow('userPrincipalName is required');
    });

    test('should throw error for missing groupId', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com'
      };

      await expect(script.invoke(params, mockContext)).rejects.toThrow('groupId is required');
    });

    test('should throw error for missing ADDRESS', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      const contextMissingAddress = {
        ...mockContext,
        environment: {}
      };

      await expect(script.invoke(params, contextMissingAddress)).rejects.toThrow('No URL specified. Provide address parameter or ADDRESS environment variable');
    });

    test('should throw error for missing OAuth2 authentication', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      const contextMissingToken = {
        ...mockContext,
        secrets: {}
      };

      await expect(script.invoke(params, contextMissingToken)).rejects.toThrow('No authentication configured');
    });

    test('should handle API error responses', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      global.fetch.mockResolvedValueOnce({
        status: 500,
        statusText: 'Internal Server Error',
        ok: false,
        text: jest.fn().mockResolvedValue('Server error occurred')
      });

      await expect(script.invoke(params, mockContext)).rejects.toThrow('Failed to add user to group: 500 Internal Server Error - Server error occurred');
    });

    test('should handle bad request with non-member error', async () => {
      const params = {
        userPrincipalName: 'invalid-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      global.fetch.mockResolvedValueOnce({
        status: 400,
        ok: false,
        text: jest.fn().mockResolvedValue('Invalid user ID')
      });

      await expect(script.invoke(params, mockContext)).rejects.toThrow('Bad request: Invalid user ID');
    });

    test('should handle network errors', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      global.fetch.mockRejectedValueOnce(new Error('Network error'));

      await expect(script.invoke(params, mockContext)).rejects.toThrow('Network error');
    });
  });

  describe('error handler', () => {
    test('should re-throw error and let framework handle retries', async () => {
      const errorObj = new Error('Rate limited: 429');
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: errorObj
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
      expect(console.error).toHaveBeenCalledWith(
        'User group assignment failed for user test-user@example.com to group 12345678-1234-1234-1234-123456789012: Rate limited: 429'
      );
    });

    test('should re-throw server errors', async () => {
      const errorObj = new Error('Server error: 502');
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: errorObj
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
    });

    test('should re-throw authentication errors', async () => {
      const errorObj = new Error('Authentication failed: 401');
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: errorObj
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
    });

    test('should re-throw any error', async () => {
      const errorObj = new Error('Unknown error occurred');
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: errorObj
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
    });
  });

  describe('halt handler', () => {
    test('should handle graceful shutdown with parameters', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        reason: 'timeout'
      };

      const result = await script.halt(params, mockContext);

      expect(result.status).toBe('halted');
      expect(result.userPrincipalName).toBe('test-user@example.com');
      expect(result.groupId).toBe('12345678-1234-1234-1234-123456789012');
      expect(result.reason).toBe('timeout');
      expect(result.halted_at).toBeDefined();
    });

    test('should handle halt without parameters', async () => {
      const params = {
        reason: 'system_shutdown'
      };

      const result = await script.halt(params, mockContext);

      expect(result.status).toBe('halted');
      expect(result.userPrincipalName).toBe('unknown');
      expect(result.groupId).toBe('unknown');
      expect(result.reason).toBe('system_shutdown');
    });
  });
});