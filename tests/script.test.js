import { jest } from '@jest/globals';
import script from '../src/script.mjs';

// Mock fetch globally
global.fetch = jest.fn();
global.URL = URL;

describe('Azure AD Add User to Group Script', () => {
  const mockContext = {
    environment: {
      AZURE_AD_TENANT_URL: 'https://graph.microsoft.com'
    },
    secrets: {
      BEARER_AUTH_TOKEN: 'test-token-123456'
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

    test('should throw error for missing AZURE_AD_TENANT_URL', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      const contextMissingTenantUrl = {
        ...mockContext,
        environment: {}
      };

      await expect(script.invoke(params, contextMissingTenantUrl)).rejects.toThrow('AZURE_AD_TENANT_URL environment variable is required');
    });

    test('should throw error for missing BEARER_AUTH_TOKEN', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      };

      const contextMissingToken = {
        ...mockContext,
        secrets: {}
      };

      await expect(script.invoke(params, contextMissingToken)).rejects.toThrow('BEARER_AUTH_TOKEN secret is required');
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
    test('should handle retryable error (429 rate limit)', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: { message: 'Rate limited: 429' }
      };

      // Mock setTimeout to avoid actual delay in tests
      jest.useFakeTimers();

      // First call for recovery attempt
      global.fetch.mockResolvedValueOnce({
        status: 204,
        ok: true
      });

      const recoveryPromise = script.error(params, mockContext);

      // Fast-forward the timer
      jest.advanceTimersByTime(5000);

      const result = await recoveryPromise;

      expect(result.status).toBe('recovered');
      expect(result.userPrincipalName).toBe('test-user@example.com');
      expect(result.groupId).toBe('12345678-1234-1234-1234-123456789012');
      expect(result.added).toBe(true);

      jest.useRealTimers();
    });

    test('should handle retryable server errors (502, 503, 504)', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: { message: 'Server error: 502' }
      };

      const result = await script.error(params, mockContext);

      expect(result.status).toBe('retry_requested');
    }, 10000); // Increase timeout for this test

    test('should not retry authentication errors (401, 403)', async () => {
      const errorObj = new Error('Authentication failed: 401');
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: errorObj
      };

      await expect(script.error(params, mockContext)).rejects.toThrow(errorObj);
    });

    test('should request retry for other errors', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: { message: 'Unknown error occurred' }
      };

      const result = await script.error(params, mockContext);

      expect(result.status).toBe('retry_requested');
    });

    test('should handle recovery failure gracefully', async () => {
      const params = {
        userPrincipalName: 'test-user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        error: { message: 'Rate limited: 429' }
      };

      jest.useFakeTimers();

      // Mock recovery attempt failure
      global.fetch.mockRejectedValueOnce(new Error('Recovery failed'));

      const recoveryPromise = script.error(params, mockContext);

      jest.advanceTimersByTime(5000);

      const result = await recoveryPromise;

      expect(result.status).toBe('retry_requested');

      jest.useRealTimers();
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