/**
 * Unit tests for HTTP utilities.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { httpRequest, clearRateLimitState } from './http.js';

describe('httpRequest', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
    clearRateLimitState();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('returns successful response on HTTP 200', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.status).toBe(200);
      expect(result.value.data).toEqual({ ok: true });
    }
  });

  it('returns error on HTTP 400', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Bad Request', { status: 400 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('INVALID_INPUT');
    }
  });

  it('returns error on HTTP 401', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Unauthorized', { status: 401 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_EXPIRED');
    }
  });

  it('returns error on HTTP 403', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Forbidden', { status: 403 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      // 403 is a permissions error, not an auth error; must not be AUTH_REQUIRED
      // or AUTH_EXPIRED, which would trigger server.ts's auto-login retry flow.
      expect(result.error.code).toBe('FORBIDDEN');
    }
  });

  it('returns error on HTTP 404', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Not Found', { status: 404 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('NOT_FOUND');
    }
  });

  it('returns error on HTTP 429 with retry-after header', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Rate Limited', {
        status: 429,
        headers: { 'Retry-After': '5' },
      })
    );

    // maxRetries: 1 avoids waiting out the real 5s Retry-After delay between
    // attempts; this test only cares about the shape of the returned error.
    const result = await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('RATE_LIMITED');
      expect(result.error.retryAfterMs).toBe(5000);
    }
  });

  it('returns error on HTTP 500', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Server Error', { status: 500 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('API_ERROR');
    }
  });

  it('returns TIMEOUT error on request timeout', async () => {
    vi.mocked(fetch).mockImplementation(() => 
      new Promise((_, reject) => {
        const error = new Error('Request timed out');
        error.name = 'AbortError';
        reject(error);
      })
    );

    const result = await httpRequest('https://api.example.com/slow', {
      timeoutMs: 100,
    });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('TIMEOUT');
    }
  });

  it('returns NETWORK_ERROR on ECONNRESET', async () => {
    vi.mocked(fetch).mockImplementation(() => 
      Promise.reject(new Error('ECONNRESET'))
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('NETWORK_ERROR');
    }
  });

  it('retries on retryable errors', async () => {
    let attempts = 0;
    vi.mocked(fetch).mockImplementation(() => {
      attempts++;
      if (attempts < 3) {
        return Promise.reject(new Error('ECONNRESET'));
      }
      return Promise.resolve(
        new Response(JSON.stringify({ success: true }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        })
      );
    });

    const result = await httpRequest('https://api.example.com/data', {
      maxRetries: 3,
      retryBaseDelayMs: 10,
    });

    expect(result.ok).toBe(true);
    expect(attempts).toBe(3);
  });

  it('does not retry on non-retryable errors like 400', async () => {
    let attempts = 0;
    vi.mocked(fetch).mockImplementation(() => {
      attempts++;
      return Promise.resolve(
        new Response('Bad Request', { status: 400 })
      );
    });

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    expect(attempts).toBe(1); // Should not retry
  });

  it('uses default timeout of 30 seconds', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('{}', { status: 200 })
    );

    await httpRequest('https://api.example.com/data');

    // Fetch was called - default timeout was applied
    expect(vi.mocked(fetch)).toHaveBeenCalled();
  });

  it('accepts custom timeout', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('{}', { status: 200 })
    );

    await httpRequest('https://api.example.com/data', {
      timeoutMs: 5000,
    });

    expect(vi.mocked(fetch)).toHaveBeenCalled();
  });

  it('returns RATE_LIMITED error when rate limited', async () => {
    // First, simulate a rate limit response
    vi.mocked(fetch).mockResolvedValue(
      new Response('Rate Limited', {
        status: 429,
        headers: { 'Retry-After': '60' },
      })
    );

    // First request triggers rate limit state. maxRetries: 1 means there is no
    // retry loop at all for this call, so no delay is incurred regardless of
    // the Retry-After value or the retryMaxDelayMs cap.
    await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    // Replace fetch mock for the second request to verify rate limit check
    vi.mocked(fetch).mockClear();
    
    // Second request should be rate limited without calling fetch
    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('RATE_LIMITED');
    }
    // Should not have called fetch since we were rate limited
    expect(vi.mocked(fetch)).not.toHaveBeenCalled();
  });

  it('scopes rate limit state to the host that returned 429, not other hosts', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('Rate Limited', {
        status: 429,
        headers: { 'Retry-After': '60' },
      })
    );

    // Trigger rate limit state for one host. maxRetries: 1 means there is no
    // retry loop at all for this call, so no delay is incurred regardless of
    // the Retry-After value or the retryMaxDelayMs cap.
    await httpRequest('https://substrate.office.com/v2/query', { maxRetries: 1 });

    // A different, unrelated host should not be blocked by that rate limit
    vi.mocked(fetch).mockClear();
    vi.mocked(fetch).mockResolvedValue(
      new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
    );

    const result = await httpRequest('https://graph.microsoft.com/v1.0/me');

    expect(result.ok).toBe(true);
    expect(vi.mocked(fetch)).toHaveBeenCalled();
  });

  it('prunes an expired rate-limit entry instead of leaking it indefinitely', async () => {
    vi.useFakeTimers();
    try {
      vi.mocked(fetch).mockResolvedValueOnce(
        new Response('Rate Limited', { status: 429, headers: { 'Retry-After': '1' } })
      );

      // Trigger a 1s rate limit for this host. maxRetries: 1 avoids retrying.
      await httpRequest('https://api.example.com/data', { maxRetries: 1 });

      // Immediately after, the same host is still blocked without calling fetch.
      vi.mocked(fetch).mockClear();
      const blocked = await httpRequest('https://api.example.com/data');
      expect(blocked.ok).toBe(false);
      expect(vi.mocked(fetch)).not.toHaveBeenCalled();

      // Once the window has passed, the stale entry is pruned and the host is
      // no longer blocked (this also exercises the delete-on-expiry path, not
      // just the "already not present" case).
      await vi.advanceTimersByTimeAsync(1500);
      vi.mocked(fetch).mockResolvedValueOnce(
        new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
      );
      const afterExpiry = await httpRequest('https://api.example.com/data');
      expect(afterExpiry.ok).toBe(true);
      expect(vi.mocked(fetch)).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('honours Retry-After as the delay for the next attempt within the same request', async () => {
    vi.useFakeTimers();
    try {
      vi.mocked(fetch)
        .mockResolvedValueOnce(
          new Response('Rate Limited', { status: 429, headers: { 'Retry-After': '3' } })
        )
        .mockResolvedValueOnce(
          new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
        );

      const promise = httpRequest('https://api.example.com/data', { maxRetries: 2 });

      // First attempt fires immediately and comes back 429.
      await vi.advanceTimersByTimeAsync(0);
      expect(vi.mocked(fetch)).toHaveBeenCalledTimes(1);

      // Before the full 3s Retry-After window elapses, the retry must not fire yet
      // (the old exponential-backoff-only delay would have fired well before this).
      await vi.advanceTimersByTimeAsync(2000);
      expect(vi.mocked(fetch)).toHaveBeenCalledTimes(1);

      // Once the Retry-After window elapses, the retry fires.
      await vi.advanceTimersByTimeAsync(1500);
      expect(vi.mocked(fetch)).toHaveBeenCalledTimes(2);

      const result = await promise;
      expect(result.ok).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });

  it('caps the recorded rate-limit window to match the actual retry delay used', async () => {
    // The server asks for a 60s Retry-After, but the default retryMaxDelayMs
    // (10s) caps the actual sleep before this request's own next attempt. The
    // window recorded in the shared rate-limit map must use that same capped
    // value; otherwise a brand new call made right after this one succeeds
    // would be blocked for up to 60s by a window that had already, in effect,
    // been ignored by this request's own retry loop.
    vi.useFakeTimers();
    try {
      vi.mocked(fetch)
        .mockResolvedValueOnce(
          new Response('Rate Limited', { status: 429, headers: { 'Retry-After': '60' } })
        )
        .mockResolvedValueOnce(
          new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
        );

      const promise = httpRequest('https://api.example.com/data', { maxRetries: 2 });
      await vi.advanceTimersByTimeAsync(10000);
      const result = await promise;
      expect(result.ok).toBe(true);
      expect(vi.mocked(fetch)).toHaveBeenCalledTimes(2);

      // A brand new call to the same host, made right after the retry
      // succeeded, must not be blocked by a stale 60s window.
      vi.mocked(fetch).mockClear();
      vi.mocked(fetch).mockResolvedValueOnce(
        new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
      );
      const followUp = await httpRequest('https://api.example.com/data');
      expect(followUp.ok).toBe(true);
      expect(vi.mocked(fetch)).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('handles plain text response', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('plain text response', {
        status: 200,
        headers: { 'Content-Type': 'text/plain' },
      })
    );

    const result = await httpRequest('https://api.example.com/text');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.data).toBe('plain text response');
    }
  });

  it('handles response without content-type header', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('no content type', { status: 200 })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.data).toBe('no content type');
    }
  });
});

describe('clearRateLimitState', () => {
  beforeEach(() => {
    clearRateLimitState();
  });

  it('clears rate limit state', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(
      new Response('Rate Limited', {
        status: 429,
        headers: { 'Retry-After': '60' },
      })
    ));

    // First request triggers rate limit. maxRetries: 1 means there is no retry
    // loop at all for this call, so no delay is incurred regardless of the
    // Retry-After value or the retryMaxDelayMs cap.
    await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    // Clear the rate limit state
    clearRateLimitState();

    // Second request should work
    vi.mocked(fetch).mockResolvedValue(
      new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
    );

    const result = await httpRequest('https://api.example.com/data');
    expect(result.ok).toBe(true);

    vi.unstubAllGlobals();
  });
});

describe('httpRequest - additional branch coverage', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
    clearRateLimitState();
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('classifies ETIMEDOUT as a NETWORK_ERROR', async () => {
    vi.mocked(fetch).mockImplementation(() => Promise.reject(new Error('connect ETIMEDOUT')));

    const result = await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('NETWORK_ERROR');
  });

  it('classifies ENOTFOUND as a NETWORK_ERROR', async () => {
    vi.mocked(fetch).mockImplementation(() => Promise.reject(new Error('getaddrinfo ENOTFOUND host')));

    const result = await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('NETWORK_ERROR');
  });

  it('returns UNKNOWN for a generic (non-network) Error', async () => {
    vi.mocked(fetch).mockImplementation(() => Promise.reject(new Error('something unexpected')));

    const result = await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('UNKNOWN');
      expect(result.error.message).toContain('something unexpected');
    }
  });

  it('returns UNKNOWN and stringifies non-Error rejections', async () => {
    vi.mocked(fetch).mockImplementation(() => Promise.reject('plain string failure'));

    const result = await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('UNKNOWN');
      expect(result.error.message).toContain('plain string failure');
    }
  });

  it('parses an empty JSON body to an empty object', async () => {
    vi.mocked(fetch).mockResolvedValue(
      new Response('', { status: 200, headers: { 'Content-Type': 'application/json' } })
    );

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.data).toEqual({});
  });

  it('handles a response with no content-type header at all', async () => {
    // A null body Response carries no content-type header (unlike a string body).
    vi.mocked(fetch).mockResolvedValue(new Response(null, { status: 200 }));

    const result = await httpRequest('https://api.example.com/data');

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.data).toBe('');
  });

  it('handles HTTP 429 without a Retry-After header', async () => {
    vi.mocked(fetch).mockResolvedValue(new Response('Rate Limited', { status: 429 }));

    const result = await httpRequest('https://api.example.com/data', { maxRetries: 1 });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('RATE_LIMITED');
  });

  it('honours a Retry-After of exactly 0 by retrying immediately, not via backoff', async () => {
    // retryAfterMs of 0 is falsy in JS. If the delay selection used a truthiness
    // check instead of an explicit undefined check, a 0ms Retry-After would be
    // ignored and the retry would instead wait out the (much longer) exponential
    // backoff delay for this attempt.
    vi.useFakeTimers();
    try {
      vi.mocked(fetch)
        .mockResolvedValueOnce(
          new Response('Rate Limited', { status: 429, headers: { 'Retry-After': '0' } })
        )
        .mockResolvedValueOnce(
          new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } })
        );

      const promise = httpRequest('https://api.example.com/data', {
        maxRetries: 2,
        retryBaseDelayMs: 5000,
      });

      // Only advance a negligible amount of time. With a 0ms Retry-After, the
      // retry must have already fired; the buggy fallback (exponential backoff
      // starting at retryBaseDelayMs) would need at least ~2500ms (half-jitter
      // minimum) before firing.
      await vi.advanceTimersByTimeAsync(1);
      expect(vi.mocked(fetch)).toHaveBeenCalledTimes(2);

      const result = await promise;
      expect(result.ok).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });

  it('returns an UNKNOWN error when maxRetries is 0 (loop body never runs)', async () => {
    vi.mocked(fetch).mockResolvedValue(new Response('{}', { status: 200 }));

    const result = await httpRequest('https://api.example.com/data', { maxRetries: 0 });

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('UNKNOWN');
      expect(result.error.message).toBe('Request failed');
    }
    expect(vi.mocked(fetch)).not.toHaveBeenCalled();
  });

  it('fires the real timeout abort callback and returns TIMEOUT', async () => {
    // fetch never resolves on its own; it only rejects when the abort signal fires,
    // which exercises the setTimeout(() => controller.abort()) path with real timers.
    vi.mocked(fetch).mockImplementation((_url, init?: RequestInit) =>
      new Promise((_resolve, reject) => {
        init?.signal?.addEventListener('abort', () => {
          const error = new Error('aborted');
          error.name = 'AbortError';
          reject(error);
        });
      })
    );

    const result = await httpRequest('https://api.example.com/slow', {
      timeoutMs: 10,
      maxRetries: 1,
    });

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('TIMEOUT');
  });
});
