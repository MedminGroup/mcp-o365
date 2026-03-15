import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import type { Config } from '../config';

// ── OAuth2 endpoints (common = any tenant + personal accounts) ────────────────

const DEVICE_CODE_URL =
  'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode';
const TOKEN_URL =
  'https://login.microsoftonline.com/common/oauth2/v2.0/token';

const SCOPES = [
  'openid',
  'profile',
  'email',
  'Calendars.ReadWrite',
  'Mail.ReadWrite',
  'Mail.Send',
  'Files.ReadWrite',
  'Sites.Read.All',
  'Contacts.ReadWrite',
  'People.Read',
  'User.Read',
  'OnlineMeetings.Read',
  'OnlineMeetingTranscript.Read.All',
  'offline_access',
].join(' ');

const GRAPH_ME_URL = 'https://graph.microsoft.com/v1.0/me';

// ── Types ─────────────────────────────────────────────────────────────────────

interface StoredAccount {
  username: string;
  name: string;
  tenantId: string;
  accessToken: string;
  refreshToken: string;
  expiresAt: number; // Unix seconds
}

interface TokenCache {
  accounts: Record<string, StoredAccount>; // keyed by username (lowercase)
}

export interface AccountSummary {
  username: string;
  name: string;
  tenantId: string;
}

export interface DeviceCodeInfo {
  verificationUri: string;
  userCode: string;
  expiresInMinutes: number;
  message: string;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

/** Decode JWT payload without verifying signature (we trust Microsoft's endpoint). */
function parseIdToken(token: string): Record<string, string> {
  try {
    const part = token.split('.')[1];
    const padded = part + '='.repeat((4 - (part.length % 4)) % 4);
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf-8')) as Record<string, string>;
  } catch {
    return {};
  }
}

function toSummary(a: StoredAccount): AccountSummary {
  return { username: a.username, name: a.name, tenantId: a.tenantId };
}

async function post(url: string, params: Record<string, string>): Promise<Record<string, unknown>> {
  const body = new URLSearchParams(params).toString();
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  return res.json() as Promise<Record<string, unknown>>;
}

// ── AccountManager ────────────────────────────────────────────────────────────

/**
 * Manages multiple signed-in Microsoft 365 accounts using direct OAuth2
 * device-code flow (no MSAL). Tokens are stored in a local JSON file.
 *
 * One Azure app registration (public client, common tenant) covers all
 * tenants and personal accounts.
 */
export class AccountManager {
  private clientId: string;
  private cachePath: string;

  /**
   * In-flight background Promise from startDeviceCodeFlow().
   * completeDeviceCodeFlow() awaits this after the user signs in.
   */
  private pendingFlow: Promise<void> | null = null;

  constructor(config: Config) {
    this.clientId = config.clientId;
    this.cachePath = config.cachePath;
  }

  // ── Token acquisition ───────────────────────────────────────────────────────

  getTokenProvider(accountHint?: string): () => Promise<string> {
    return () => this.getToken(accountHint);
  }

  private async getToken(accountHint?: string): Promise<string> {
    const cache = this.loadCache();
    const all = Object.values(cache.accounts);

    if (all.length === 0) {
      throw new Error(
        'No accounts signed in. Use accounts_add to sign in first.',
      );
    }

    let account: StoredAccount | undefined;

    if (accountHint) {
      const lower = accountHint.toLowerCase();
      account = all.find(
        (a) =>
          a.username.toLowerCase() === lower ||
          a.username.toLowerCase().includes(lower) ||
          a.name.toLowerCase().includes(lower),
      );
      if (!account) {
        const names = all.map((a) => a.username).join(', ');
        throw new Error(`No account matching "${accountHint}". Signed-in: ${names}`);
      }
    } else {
      account = all[0];
    }

    // Refresh if < 5 minutes to expiry
    const nowSecs = Math.floor(Date.now() / 1000);
    if (account.expiresAt - nowSecs < 300) {
      account = await this.refreshAccessToken(account, cache);
    }

    return account.accessToken;
  }

  private async refreshAccessToken(
    account: StoredAccount,
    cache: TokenCache,
  ): Promise<StoredAccount> {
    const data = await post(TOKEN_URL, {
      grant_type: 'refresh_token',
      refresh_token: account.refreshToken,
      client_id: this.clientId,
      scope: SCOPES,
    });

    if (!data.access_token) {
      throw new Error(
        `Token refresh failed: ${String(data.error_description ?? data.error ?? 'unknown')}`,
      );
    }

    const updated: StoredAccount = {
      ...account,
      accessToken: String(data.access_token),
      refreshToken: data.refresh_token ? String(data.refresh_token) : account.refreshToken,
      expiresAt: Math.floor(Date.now() / 1000) + Number(data.expires_in ?? 3600),
    };

    cache.accounts[account.username.toLowerCase()] = updated;
    this.saveCache(cache);
    return updated;
  }

  // ── Two-step device-code flow ───────────────────────────────────────────────

  /**
   * STEP 1 — Called by accounts_add tool.
   *
   * Makes a single HTTP POST to Microsoft to get the device code,
   * resolves immediately with the URL + code, then continues polling
   * in the background. Call completeDeviceCodeFlow() after sign-in.
   */
  startDeviceCodeFlow(): Promise<DeviceCodeInfo> {
    if (this.pendingFlow) {
      return Promise.reject(
        new Error(
          'A sign-in is already in progress. ' +
            'Complete it with accounts_complete, or wait for it to expire.',
        ),
      );
    }

    return new Promise<DeviceCodeInfo>((resolveOuter, rejectOuter) => {
      // Reject with a clear error if the first HTTP call takes > 15 s
      let resolved = false;
      const timeout = setTimeout(() => {
        if (!resolved) {
          this.pendingFlow = null;
          rejectOuter(
            new Error(
              `No response from Microsoft after 15 s. ` +
                `Check AZURE_CLIENT_ID (currently: ${this.clientId}) ` +
                `and that login.microsoftonline.com is reachable.`,
            ),
          );
        }
      }, 15_000);

      const run = async (): Promise<void> => {
        // ── Step 1: request a device code ──────────────────────────────────
        const dc = await post(DEVICE_CODE_URL, {
          client_id: this.clientId,
          scope: SCOPES,
        });

        clearTimeout(timeout);

        if (!dc.device_code) {
          throw new Error(
            `Microsoft refused to issue a device code: ` +
              String(dc.error_description ?? dc.error ?? JSON.stringify(dc)),
          );
        }

        const interval = Number(dc.interval ?? 5);
        const expiresIn = Number(dc.expires_in ?? 900);

        resolved = true;
        resolveOuter({
          verificationUri: String(dc.verification_uri),
          userCode: String(dc.user_code),
          expiresInMinutes: Math.floor(expiresIn / 60),
          message: String(dc.message ?? ''),
        });

        // ── Step 2: poll until the user signs in ────────────────────────────
        await this.pollForToken(String(dc.device_code), interval);
      };

      this.pendingFlow = run()
        .then(() => {
          this.pendingFlow = null;
        })
        .catch((err: Error) => {
          this.pendingFlow = null;
          if (!resolved) {
            clearTimeout(timeout);
            rejectOuter(err);
          }
          // If already resolved (user saw URL), errors during polling are
          // surfaced when accounts_complete is called.
        });
    });
  }

  private async pollForToken(deviceCode: string, intervalSecs: number): Promise<void> {
    let interval = intervalSecs;

    while (true) {
      await sleep(interval * 1000);

      const data = await post(TOKEN_URL, {
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        device_code: deviceCode,
        client_id: this.clientId,
      });

      if (data.access_token) {
        const accessToken = String(data.access_token);

        // Parse identity from the id_token JWT
        const claims = data.id_token ? parseIdToken(String(data.id_token)) : {};
        let username = claims.preferred_username ?? claims.upn ?? claims.email ?? '';
        let name = claims.name ?? '';
        let tenantId = claims.tid ?? 'common';

        // Fallback: call /me if id_token didn't give us a username
        if (!username) {
          try {
            const me = await fetch(GRAPH_ME_URL, {
              headers: { Authorization: `Bearer ${accessToken}` },
            }).then((r) => r.json()) as Record<string, string>;
            username = me.userPrincipalName ?? me.mail ?? me.id ?? 'unknown';
            name = me.displayName ?? username;
          } catch {
            username = 'unknown';
            name = 'unknown';
          }
        }
        if (!name) name = username;

        const account: StoredAccount = {
          username,
          name,
          tenantId,
          accessToken,
          refreshToken: data.refresh_token ? String(data.refresh_token) : '',
          expiresAt: Math.floor(Date.now() / 1000) + Number(data.expires_in ?? 3600),
        };

        const cache = this.loadCache();
        // Remove stale 'unknown' entry if present
        delete cache.accounts['unknown'];
        cache.accounts[username.toLowerCase()] = account;
        this.saveCache(cache);
        return;
      }

      const error = String(data.error ?? '');
      if (error === 'authorization_pending') continue;
      if (error === 'slow_down') { interval += 5; continue; }
      if (error === 'expired_token') {
        throw new Error('Device code expired before sign-in completed. Call accounts_add again.');
      }
      throw new Error(String(data.error_description ?? data.error ?? 'Polling error'));
    }
  }

  /**
   * STEP 2 — Called by accounts_complete tool after the user has signed in.
   * Awaits the background polling promise and returns the new account.
   */
  async completeDeviceCodeFlow(): Promise<AccountSummary> {
    if (!this.pendingFlow) {
      // May have already completed (fast sign-in)
      const all = Object.values(this.loadCache().accounts);
      if (all.length > 0) return toSummary(all[all.length - 1]);
      throw new Error('No sign-in in progress. Call accounts_add first.');
    }
    await this.pendingFlow;
    const all = Object.values(this.loadCache().accounts);
    return toSummary(all[all.length - 1]);
  }

  // ── Account management ──────────────────────────────────────────────────────

  async listAccounts(): Promise<AccountSummary[]> {
    return Object.values(this.loadCache().accounts).map(toSummary);
  }

  async removeAccount(hint: string): Promise<AccountSummary> {
    const cache = this.loadCache();
    const all = Object.values(cache.accounts);
    const lower = hint.toLowerCase();
    const target = all.find(
      (a) =>
        a.username.toLowerCase() === lower ||
        a.username.toLowerCase().includes(lower) ||
        a.name.toLowerCase().includes(lower),
    );
    if (!target) {
      const names = all.map((a) => a.username).join(', ') || '(none)';
      throw new Error(`No account matching "${hint}". Available: ${names}`);
    }
    delete cache.accounts[target.username.toLowerCase()];
    this.saveCache(cache);
    return toSummary(target);
  }

  // ── Cache I/O ─────────────────────────────────────────────────────────────

  private loadCache(): TokenCache {
    if (!existsSync(this.cachePath)) return { accounts: {} };
    try {
      return JSON.parse(readFileSync(this.cachePath, 'utf-8')) as TokenCache;
    } catch {
      return { accounts: {} };
    }
  }

  private saveCache(cache: TokenCache): void {
    try {
      writeFileSync(this.cachePath, JSON.stringify(cache, null, 2), {
        mode: 0o600,
        encoding: 'utf-8',
      });
    } catch (err) {
      process.stderr.write(`Warning: could not save token cache: ${err}\n`);
    }
  }
}
