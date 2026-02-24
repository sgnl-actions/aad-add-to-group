/* global beforeAll, afterAll */
import { jest } from '@jest/globals';
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'fs';
import { request } from 'https';
import script from '../src/script.mjs';

const FIXTURES_DIR = '__recordings__';
const FIXTURE_FILE = `${FIXTURES_DIR}/aad-add-user-to-group.json`;
const IS_RECORDING = process.env.RECORD_MODE === 'true';

function loadFixtures() {
  if (existsSync(FIXTURE_FILE)) {
    return JSON.parse(readFileSync(FIXTURE_FILE, 'utf-8'));
  }
  return {};
}

function saveFixtures(fixtures) {
  if (!existsSync(FIXTURES_DIR)) mkdirSync(FIXTURES_DIR, { recursive: true });
  writeFileSync(FIXTURE_FILE, JSON.stringify(fixtures, null, 2));
}

function httpsRequest(url, options) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const body = options.body;
    const reqOptions = {
      hostname: parsed.hostname,
      path: parsed.pathname + parsed.search,
      method: options.method || 'GET',
      headers: {
        ...options.headers,
        ...(body ? { 'Content-Length': Buffer.byteLength(body) } : {})
      }
    };
    const req = request(reqOptions, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        const isJson = res.headers['content-type']?.includes('application/json');
        const parsedBody = isJson ? JSON.parse(data) : data;
        resolve({
          status: res.statusCode,
          ok: res.statusCode >= 200 && res.statusCode < 300,
          statusText: res.statusMessage,
          body: parsedBody
        });
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

function makeRecordReplayFetch(fixtures, key) {
  return async (url, options) => {
    if (IS_RECORDING) {
      // Always hit the real API and overwrite the fixture
      const res = await httpsRequest(url, options || {});
      fixtures[key] = { status: res.status, ok: res.ok, statusText: res.statusText, body: res.body };
      return {
        ok: res.ok, status: res.status, statusText: res.statusText,
        json: async () => res.body,
        text: async () => (typeof res.body === 'string' ? res.body : JSON.stringify(res.body ?? ''))
      };
    }

    // Replay mode: use saved fixture
    const fixture = fixtures[key];
    if (!fixture) throw new Error(`No fixture for "${key}". Run with RECORD_MODE=true first.`);
    return {
      ok: fixture.ok, status: fixture.status, statusText: fixture.statusText,
      json: async () => fixture.body,
      text: async () => (typeof fixture.body === 'string' ? fixture.body : JSON.stringify(fixture.body ?? ''))
    };
  };
}

describe('AAD Add User to Group - Record & Replay', () => {
  let fixtures = {};

  beforeAll(async () => {
    fixtures = loadFixtures();

    // In record mode, remove the user from the group before recording starts.
    // This ensures test 1 always starts with a clean state (user not in group).
    // Cleanup at the START rather than end avoids Azure propagation delays
    // that occur when running record mode back to back immediately.
    if (IS_RECORDING) {
      try {
        const tokenRes = await httpsRequest(
          process.env.AZURE_TOKEN_URL,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: `client_id=${process.env.AZURE_CLIENT_ID}&client_secret=${process.env.AZURE_CLIENT_SECRET}&scope=https://graph.microsoft.com/.default&grant_type=client_credentials`
          }
        );
        const token = tokenRes.body.access_token;

        const userRes = await httpsRequest(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.AZURE_TEST_UPN)}`,
          { method: 'GET', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } }
        );
        const userId = userRes.body.id;

        await httpsRequest(
          `https://graph.microsoft.com/v1.0/groups/${process.env.AZURE_GROUP_ID}/members/${userId}/$ref`,
          { method: 'DELETE', headers: { 'Authorization': `Bearer ${token}` } }
        );

        console.log(`Pre-test cleanup: removed ${process.env.AZURE_TEST_UPN} from group`);
      } catch (e) {
        // User may not be in group — that's fine, proceed with recording
        console.log(`Pre-test cleanup: user not in group or already removed (${e.message})`);
      }
    }
  });

  afterAll(async () => {
    if (IS_RECORDING) saveFixtures(fixtures);
  });

  beforeEach(() => {
    fetch.mockClear();
    jest.spyOn(console, 'log').mockImplementation(() => {});
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  // Fallback values ensure createAuthHeaders proceeds to call fetch in replay mode,
  // where env vars are not set. The token fetch is intercepted by the mock anyway.
  const context = {
    environment: {
      ADDRESS: 'https://graph.microsoft.com',
      OAUTH2_CLIENT_CREDENTIALS_TOKEN_URL: process.env.AZURE_TOKEN_URL || 'https://login.microsoftonline.com/test-tenant/oauth2/v2.0/token',
      OAUTH2_CLIENT_CREDENTIALS_CLIENT_ID: process.env.AZURE_CLIENT_ID || 'test-client-id',
      OAUTH2_CLIENT_CREDENTIALS_SCOPE: 'https://graph.microsoft.com/.default'
    },
    secrets: {
      OAUTH2_CLIENT_CREDENTIALS_CLIENT_SECRET: process.env.AZURE_CLIENT_SECRET || 'test-client-secret'
    },
    outputs: {}
  };

  // For synthetic error tests, bypass OAuth2 token fetch entirely
  const syntheticContext = {
    environment: { ADDRESS: 'https://graph.microsoft.com' },
    secrets: { BEARER_AUTH_TOKEN: 'fake-bearer-token-for-testing' },
    outputs: {}
  };

  const params = {
    userPrincipalName: process.env.AZURE_TEST_UPN || 'testuser@yourtenant.onmicrosoft.com',
    groupId: process.env.AZURE_GROUP_ID || 'test-group-id'
  };

  // IDEMPOTENCY: This action IS idempotent.
  // First call adds the user (204). Second call finds user already in group (400
  // with "already a member") which the script handles as a success with added:false.
  // Both calls return status:'success' — same end state.
  // Synthetic fixtures for error scenarios that can't be triggered with valid credentials.
  // These are injected directly as mock responses, bypassing makeRecordReplayFetch entirely.
  test('should add user to group successfully on first call', async () => {
    // createAuthHeaders fetches OAuth2 token first, then the Graph API call
    fetch
      .mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-oauth-token'))
      .mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-add-user'));

    const result = await script.invoke(params, context);

    expect(result.status).toBe('success');
    expect(result.added).toBe(true);
    expect(result.userPrincipalName).toBe(params.userPrincipalName);
    expect(result.groupId).toBe(params.groupId);
    expect(fetch).toHaveBeenCalledTimes(2);
  });

  test('should be idempotent - second call succeeds when user already in group', async () => {
    // Second call: user already a member — script returns success with added:false
    if (IS_RECORDING) {
      fixtures['aad-oauth-token-2'] = fixtures['aad-oauth-token'];
      fixtures['aad-add-user-already-member'] = {
        status: 400, ok: false, statusText: 'Bad Request',
        body: JSON.stringify({
          error: {
            code: 'Request_BadRequest',
            message: "One or more added object references already exist for the following modified properties: 'members'."
          }
        })
      };
    }

    fetch
      .mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-oauth-token-2'))
      .mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-add-user-already-member'));

    const result = await script.invoke(params, context);

    // Still succeeds — idempotent behavior
    expect(result.status).toBe('success');
    expect(result.added).toBe(false);
    expect(result.message).toMatch(/already a member/i);
  });

  test('should handle group not found', async () => {
    if (IS_RECORDING) {
      fixtures['aad-group-not-found'] = {
        status: 404, ok: false, statusText: 'Not Found',
        body: { error: { code: 'Request_ResourceNotFound', message: 'Resource not found' } }
      };
    }

    fetch.mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-group-not-found'));

    await expect(script.invoke(params, syntheticContext))
      .rejects.toThrow(/Failed to add user to group/);
  });

  test('should handle user not found', async () => {
    if (IS_RECORDING) {
      fixtures['aad-user-not-found'] = {
        status: 404, ok: false, statusText: 'Not Found',
        body: { error: { code: 'Request_ResourceNotFound', message: 'Resource not found' } }
      };
    }

    fetch.mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-user-not-found'));

    await expect(script.invoke(params, syntheticContext))
      .rejects.toThrow(/Failed to add user to group/);
  });

  test('should handle unauthorized (invalid token)', async () => {
    if (IS_RECORDING) {
      fixtures['aad-unauthorized'] = {
        status: 401, ok: false, statusText: 'Unauthorized',
        body: { error: { code: 'InvalidAuthenticationToken', message: 'Access token is invalid' } }
      };
    }

    fetch.mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-unauthorized'));

    await expect(script.invoke(params, syntheticContext))
      .rejects.toThrow(/Failed to add user to group/);
  });

  test('should handle insufficient permissions', async () => {
    if (IS_RECORDING) {
      fixtures['aad-forbidden'] = {
        status: 403, ok: false, statusText: 'Forbidden',
        body: { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } }
      };
    }

    fetch.mockImplementationOnce(makeRecordReplayFetch(fixtures, 'aad-forbidden'));

    await expect(script.invoke(params, syntheticContext))
      .rejects.toThrow(/Failed to add user to group/);
  });

  test('should handle missing auth token', async () => {
    await expect(script.invoke(params, {
      environment: { ADDRESS: 'https://graph.microsoft.com' },
      secrets: {},
      outputs: {}
    })).rejects.toThrow(/No authentication configured/);

    expect(fetch).not.toHaveBeenCalled();
  });
});