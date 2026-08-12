/**
 * Graph tokens for a headless run.
 *
 * The taskpane gets a token from the Outlook host with no stored credential at
 * all (Nested App Authentication). A cron job has no host and no browser, so it
 * needs a credential that survives between runs. This is the single real cost of
 * the routine over the add-in, and the whole file exists to make that cost as
 * small and as visible as possible.
 *
 * Delegated refresh token, not app-only client credentials. Two reasons:
 *
 *  - Scope. `Mail.ReadWrite` as an *application* permission is admin-consented
 *    and covers every mailbox in the tenant until an Exchange
 *    ApplicationAccessPolicy narrows it. A delegated refresh token can only ever
 *    reach the one mailbox that consented, and no admin has to approve anything.
 *  - Reuse. Delegated tokens make `/me/...` mean her mailbox, so every path in
 *    `src/graph.ts` works untouched. App-only would need each of them rewritten
 *    to `/users/{upn}/...`.
 *
 * The app registration for this must be a PUBLIC CLIENT (Mobile & desktop
 * platform, "Allow public client flows" enabled) - NOT the SPA registration the
 * add-in uses. Refresh tokens issued to a single-page app are deliberately
 * capped at 24 hours by Entra, which is useless for a nightly job; public-client
 * refresh tokens last 90 days and the window rolls forward on every use.
 */

const AUTHORITY = 'https://login.microsoftonline.com';

/**
 * `offline_access` is what makes a refresh token appear at all. The rest mirror
 * the add-in's delegated permissions exactly, so the routine can do nothing the
 * taskpane couldn't already do.
 */
export const REQUIRED_SCOPES = [
  'Mail.ReadWrite',
  'MailboxSettings.ReadWrite',
  // Read-only, and only the calendar. The weekly digest reports unanswered
  // invitations and the week ahead; nothing in this project writes an event or
  // sends an RSVP.
  'Calendars.Read',
  'User.Read',
] as const;

export const SCOPES = [
  'offline_access',
  ...REQUIRED_SCOPES.map((s) => `https://graph.microsoft.com/${s}`),
].join(' ');

/**
 * Which required scopes are missing from what Entra actually granted.
 *
 * Compared on the short name because Entra reports them inconsistently -
 * sometimes bare, sometimes fully qualified with the resource URI.
 */
export function missingScopes(grantedScopes: string): string[] {
  const granted = new Set(
    grantedScopes
      .split(/\s+/)
      .filter(Boolean)
      .map((s) => (s.includes('/') ? (s.split('/').pop() as string) : s))
      .map((s) => s.toLowerCase()),
  );
  return REQUIRED_SCOPES.filter((s) => !granted.has(s.toLowerCase()));
}

export interface TokenSet {
  accessToken: string;
  /** Entra rotates this on every exchange. Persisting the new one keeps the 90-day window rolling. */
  refreshToken: string;
  expiresAt: number;
  /**
   * Scopes actually granted, space-separated, as reported by Entra.
   *
   * Worth carrying because it is the only reliable way to catch a permission
   * that was never added to the app registration: sign-in succeeds either way,
   * and the omission would otherwise surface much later as an opaque Graph 403.
   */
  grantedScopes: string;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope?: string;
}

function config() {
  const clientId = process.env.STEWARD_CLIENT_ID;
  const tenant = process.env.STEWARD_TENANT ?? 'organizations';
  if (!clientId) {
    throw new Error(
      'STEWARD_CLIENT_ID is not set. See README.md - this is the public-client app registration, not the add-in SPA one.',
    );
  }
  return { clientId, tenant };
}

/**
 * Fail on missing configuration before anything else is attempted, so a missing
 * client id doesn't surface as the misleading "no refresh token" further down.
 */
export function assertConfigured(): void {
  config();
}

async function tokenEndpoint(body: Record<string, string>): Promise<TokenResponse> {
  const { tenant } = config();
  const res = await fetch(`${AUTHORITY}/${tenant}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(body),
  });

  const payload = (await res.json()) as TokenResponse & {
    error?: string;
    error_description?: string;
  };

  if (!res.ok) {
    // invalid_grant is the one worth naming: it means the refresh token expired
    // or was revoked, and the fix is a human re-running `login`, not a retry.
    const hint =
      payload.error === 'invalid_grant'
        ? ' The stored refresh token is no longer valid - run `npm run login` again.'
        : '';
    throw new Error(
      `Token request failed (${res.status}): ${payload.error_description ?? payload.error ?? 'unknown'}.${hint}`,
    );
  }
  return payload;
}

export async function redeemRefreshToken(refreshToken: string): Promise<TokenSet> {
  const { clientId } = config();
  const res = await tokenEndpoint({
    client_id: clientId,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: SCOPES,
  });

  return {
    accessToken: res.access_token,
    // Entra normally returns a fresh one; falling back to the old token is
    // correct because the previous token stays valid until its own window ends.
    refreshToken: res.refresh_token ?? refreshToken,
    expiresAt: Date.now() + res.expires_in * 1000,
    grantedScopes: res.scope ?? '',
  };
}

// ---------------------------------------------------------------------------
// One-time interactive consent
// ---------------------------------------------------------------------------

export interface DeviceCode {
  userCode: string;
  verificationUri: string;
  message: string;
  deviceCode: string;
  interval: number;
  expiresIn: number;
}

/**
 * Device code flow, run once by a human to mint the first refresh token.
 *
 * Chosen over an authorization-code redirect because it needs no local web
 * server and no redirect URI reachable from the machine doing the setup - the
 * person just opens a URL and types a code.
 */
export async function startDeviceCode(): Promise<DeviceCode> {
  const { clientId, tenant } = config();
  const res = await fetch(`${AUTHORITY}/${tenant}/oauth2/v2.0/devicecode`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: clientId, scope: SCOPES }),
  });
  if (!res.ok) {
    const detail = await res.text();

    // Two very different causes reach this branch, and conflating them sends
    // people to the wrong portal page. AADSTS50059 means Entra could not work
    // out which directory to ask - a tenant problem. Everything else here is
    // almost always the public-client setting.
    const cause = /AADSTS50059|tenant-identifying/i.test(detail)
      ? 'Set STEWARD_TENANT to the Directory (tenant) ID from the app registration\'s Overview page.'
      : 'Check that the app registration has "Allow public client flows" set to Yes on its Authentication page.';

    throw new Error(`Could not start sign-in (${res.status}). ${cause}\n\n${detail}`);
  }
  const body = (await res.json()) as {
    user_code: string;
    verification_uri: string;
    message: string;
    device_code: string;
    interval: number;
    expires_in: number;
  };
  return {
    userCode: body.user_code,
    verificationUri: body.verification_uri,
    message: body.message,
    deviceCode: body.device_code,
    interval: body.interval,
    expiresIn: body.expires_in,
  };
}

export async function pollDeviceCode(code: DeviceCode): Promise<TokenSet> {
  const { clientId } = config();
  const deadline = Date.now() + code.expiresIn * 1000;
  // Backed off on `slow_down`, per the device-code spec. Polling on at the same
  // cadence just earns another slow_down and burns the whole window.
  let intervalSeconds = code.interval;

  for (;;) {
    if (Date.now() > deadline) throw new Error('Device code expired before sign-in completed.');
    await new Promise((r) => setTimeout(r, intervalSeconds * 1000));

    const { tenant } = config();
    const res = await fetch(`${AUTHORITY}/${tenant}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: clientId,
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        device_code: code.deviceCode,
      }),
    });
    const body = (await res.json()) as TokenResponse & { error?: string };

    if (res.ok) {
      if (!body.refresh_token) {
        throw new Error(
          'Sign-in succeeded but no refresh token was issued. The app registration is probably still configured as a single-page application; add a "Mobile and desktop applications" platform instead.',
        );
      }
      return {
        accessToken: body.access_token,
        refreshToken: body.refresh_token,
        expiresAt: Date.now() + body.expires_in * 1000,
        grantedScopes: body.scope ?? '',
      };
    }

    // These two are the flow working as designed, not failures.
    if (body.error === 'slow_down') {
      intervalSeconds += 5;
      continue;
    }
    if (body.error === 'authorization_pending') continue;
    throw new Error(`Device code sign-in failed: ${body.error ?? res.status}`);
  }
}
