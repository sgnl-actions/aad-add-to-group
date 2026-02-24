import { jest } from '@jest/globals';
import script from '../src/script.mjs';
import { SGNL_USER_AGENT } from '@sgnl-actions/utils';

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
            'Accept': 'application/json',
            'User-Agent': SGNL_USER_AGENT
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

  describe('invoke handler - idempotency', () => {
    test('should succeed on first call with added:true', async () => {
      global.fetch.mockResolvedValueOnce({ status: 204, ok: true });

      const result = await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext);

      expect(result.status).toBe('success');
      expect(result.added).toBe(true);
    });

    test('should succeed on second call with added:false (already a member)', async () => {
      global.fetch.mockResolvedValueOnce({
        status: 400, ok: false,
        text: jest.fn().mockResolvedValue('User is already a member of the group')
      });

      const result = await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext);

      expect(result.status).toBe('success');
      expect(result.added).toBe(false);
    });

    test('should produce same end state on repeated calls', async () => {
      global.fetch
        .mockResolvedValueOnce({ status: 204, ok: true })
        .mockResolvedValueOnce({
          status: 400, ok: false,
          text: jest.fn().mockResolvedValue('User is already a member of the group')
        });

      const r1 = await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext);

      const r2 = await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext);

      // Both succeed — idempotent
      expect(r1.status).toBe('success');
      expect(r2.status).toBe('success');
      expect(r1.userPrincipalName).toBe(r2.userPrincipalName);
      expect(r1.groupId).toBe(r2.groupId);
    });

    test('should handle real Azure already-member error message', async () => {
      global.fetch.mockResolvedValueOnce({
        status: 400, ok: false,
        text: jest.fn().mockResolvedValue(JSON.stringify({
          error: {
            code: 'Request_BadRequest',
            message: "One or more added object references already exist for the following modified properties: 'members'."
          }
        }))
      });

      const result = await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext);

      expect(result.status).toBe('success');
      expect(result.added).toBe(false);
    });
  });

  describe('invoke handler - input validation', () => {
    test('should throw when userPrincipalName is missing', async () => {
      await expect(script.invoke({
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext)).rejects.toThrow('userPrincipalName parameter is required and cannot be empty');
      expect(global.fetch).not.toHaveBeenCalled();
    });

    test('should throw when groupId is missing', async () => {
      await expect(script.invoke({
        userPrincipalName: 'user@example.com'
      }, mockContext)).rejects.toThrow('groupId parameter is required and cannot be empty');
      expect(global.fetch).not.toHaveBeenCalled();
    });

    test('should throw when auth token is missing', async () => {
      await expect(script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, { environment: { ADDRESS: 'https://graph.microsoft.com' }, secrets: {} }))
        .rejects.toThrow(/No authentication configured/);

      expect(global.fetch).not.toHaveBeenCalled();
    });
  });

  describe('invoke handler - request construction', () => {
    test('should use custom address from params over environment ADDRESS', async () => {
      global.fetch.mockResolvedValueOnce({ status: 204, ok: true });

      await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012',
        address: 'https://custom-proxy.example.com'
      }, mockContext);

      expect(global.fetch).toHaveBeenCalledWith(
        expect.stringContaining('https://custom-proxy.example.com'),
        expect.any(Object)
      );
    });

    test('should include User-Agent header', async () => {
      global.fetch.mockResolvedValueOnce({ status: 204, ok: true });

      await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext);

      expect(global.fetch).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({ 'User-Agent': SGNL_USER_AGENT })
        })
      );
    });

    test('should strip trailing slash from base URL', async () => {
      global.fetch.mockResolvedValueOnce({ status: 204, ok: true });

      await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, {
        ...mockContext,
        environment: { ADDRESS: 'https://graph.microsoft.com/' } // trailing slash
      });

      expect(global.fetch).toHaveBeenCalledWith(
        expect.not.stringContaining('//v1.0'),
        expect.any(Object)
      );
    });

    test('should use OData reference format in request body', async () => {
      global.fetch.mockResolvedValueOnce({ status: 204, ok: true });

      await script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext);

      const body = JSON.parse(global.fetch.mock.calls[0][1].body);
      expect(body['@odata.id']).toMatch(/^https:\/\/graph\.microsoft\.com\/v1\.0\/users\//);
    });
  });

  describe('invoke handler - error responses', () => {
    test('should throw on 401 Unauthorized', async () => {
      global.fetch.mockResolvedValueOnce({
        status: 401, statusText: 'Unauthorized', ok: false,
        text: jest.fn().mockResolvedValue('{"error":{"code":"InvalidAuthenticationToken"}}')
      });

      await expect(script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext)).rejects.toThrow(/401 Unauthorized/);
    });

    test('should throw on 403 Forbidden', async () => {
      global.fetch.mockResolvedValueOnce({
        status: 403, statusText: 'Forbidden', ok: false,
        text: jest.fn().mockResolvedValue('{"error":{"code":"Authorization_RequestDenied"}}')
      });

      await expect(script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext)).rejects.toThrow(/403 Forbidden/);
    });

    test('should throw on 404 Not Found', async () => {
      global.fetch.mockResolvedValueOnce({
        status: 404, statusText: 'Not Found', ok: false,
        text: jest.fn().mockResolvedValue('{"error":{"code":"Request_ResourceNotFound"}}')
      });

      await expect(script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext)).rejects.toThrow(/404 Not Found/);
    });

    test('should throw on 429 Too Many Requests', async () => {
      global.fetch.mockResolvedValueOnce({
        status: 429, statusText: 'Too Many Requests', ok: false,
        text: jest.fn().mockResolvedValue('{"error":{"code":"TooManyRequests"}}')
      });

      await expect(script.invoke({
        userPrincipalName: 'user@example.com',
        groupId: '12345678-1234-1234-1234-123456789012'
      }, mockContext)).rejects.toThrow(/429/);
    });
  });

  describe('halt handler - edge cases', () => {
    test('should handle halt with no params at all', async () => {
      const result = await script.halt({}, mockContext);

      expect(result.status).toBe('halted');
      expect(result.userPrincipalName).toBe('unknown');
      expect(result.groupId).toBe('unknown');
      expect(result.reason).toBeUndefined();
      expect(result.halted_at).toBeDefined();
    });

    test('should include ISO timestamp in halted_at', async () => {
      const result = await script.halt({ reason: 'test' }, mockContext);
      expect(new Date(result.halted_at).toISOString()).toBe(result.halted_at);
    });
  });
});