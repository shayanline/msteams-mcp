/**
 * HTTP utilities with retry, timeout, and error handling.
 */

import { 
  ErrorCode, 
  createError, 
  classifyHttpError, 
  extractRetryAfter,
  type McpError,
} from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';

/** Options for HTTP requests. */
export interface HttpOptions extends Omit<RequestInit, 'signal'> {
  /** Timeout in milliseconds (default: 30000). */
  timeoutMs?: number;
  /** Maximum retry attempts (default: 3). */
  maxRetries?: number;
  /** Base delay for exponential backoff in ms (default: 1000). */
  retryBaseDelayMs?: number;
  /** Maximum delay between retries in ms (default: 10000). */
  retryMaxDelayMs?: number;
}

/** Successful HTTP response. */
export interface HttpResponse<T = unknown> {
  status: number;
  headers: Headers;
  data: T;
}

/** Rate limit state tracking, keyed by host so one throttled API doesn't block the rest. */
const rateLimitedUntil = new Map<string, number>();

/**
 * Returns the host (hostname plus port, if any) to key rate-limit state by.
 * Falls back to the full URL if it can't be parsed as an absolute URL, so an
 * invalid input never crashes the request path.
 */
function getRateLimitKey(url: string): string {
  try {
    return new URL(url).host;
  } catch {
    return url;
  }
}

/**
 * Returns the active rate-limit expiry for a key, if any, relative to `now`.
 * Deletes the entry first if it has already passed, so a long-running process
 * that talks to many distinct hosts doesn't accumulate stale keys
 * indefinitely. Taking `now` as a parameter (rather than calling Date.now()
 * again here) guarantees the caller's own "time remaining" calculation is
 * against the exact same instant used to decide the entry is still active,
 * so it can never come back zero or negative.
 */
function getActiveRateLimitExpiry(key: string, now: number): number | undefined {
  const expiry = rateLimitedUntil.get(key);
  if (expiry === undefined) return undefined;
  if (now >= expiry) {
    rateLimitedUntil.delete(key);
    return undefined;
  }
  return expiry;
}

/**
 * Makes an HTTP request with timeout, retry, and error handling.
 */
export async function httpRequest<T = unknown>(
  url: string,
  options: HttpOptions = {}
): Promise<Result<HttpResponse<T>>> {
  const {
    timeoutMs = 30000,
    maxRetries = 3,
    retryBaseDelayMs = 1000,
    retryMaxDelayMs = 10000,
    ...fetchOptions
  } = options;

  const rateLimitKey = getRateLimitKey(url);

  // Check rate limit state. waitMs is computed from the same `now` used to
  // decide the entry is still active, so it can never come back zero/negative
  // (which could otherwise happen if the entry expired between the two calls).
  const now = Date.now();
  const rateLimitExpiry = getActiveRateLimitExpiry(rateLimitKey, now);
  if (rateLimitExpiry !== undefined) {
    const waitMs = rateLimitExpiry - now;
    return err(createError(
      ErrorCode.RATE_LIMITED,
      `Rate limited. Retry after ${Math.ceil(waitMs / 1000)}s`,
      { retryable: true, retryAfterMs: waitMs }
    ));
  }

  let lastError: McpError | null = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const result = await fetchWithTimeout<T>(url, fetchOptions, timeoutMs);
      
      if (result.ok) {
        return result;
      }

      lastError = result.error;

      // Handle rate limiting. Check for undefined rather than truthiness so a
      // valid "Retry-After: 0" is still recorded instead of being ignored.
      if (result.error.code === ErrorCode.RATE_LIMITED && result.error.retryAfterMs !== undefined) {
        rateLimitedUntil.set(rateLimitKey, Date.now() + result.error.retryAfterMs);
      }

      // Don't retry non-retryable errors
      if (!result.error.retryable) {
        return result;
      }

      // Don't retry on last attempt
      if (attempt === maxRetries) {
        return result;
      }

      // Honour the server's Retry-After for the next attempt of this same request
      // (checking for undefined rather than truthiness so a valid 0ms value is
      // still honoured instead of falling through to backoff); otherwise fall
      // back to exponential backoff with jitter to avoid thundering herd.
      const delay = result.error.retryAfterMs !== undefined
        ? Math.min(result.error.retryAfterMs, retryMaxDelayMs)
        : Math.min(retryBaseDelayMs * Math.pow(2, attempt - 1), retryMaxDelayMs) * (0.5 + Math.random() * 0.5);

      await sleep(delay);
      
    } catch (error) {
      // Unexpected errors
      lastError = createError(
        ErrorCode.UNKNOWN,
        error instanceof Error ? error.message : String(error),
        { retryable: false }
      );
      
      if (attempt === maxRetries) {
        return err(lastError);
      }
    }
  }

  return err(lastError ?? createError(ErrorCode.UNKNOWN, 'Request failed'));
}

/**
 * Makes a single fetch request with timeout.
 */
async function fetchWithTimeout<T>(
  url: string,
  options: RequestInit,
  timeoutMs: number
): Promise<Result<HttpResponse<T>>> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    // Handle error responses
    if (!response.ok) {
      const retryAfterMs = extractRetryAfter(response.headers);
      const errorText = await response.text().catch(() => '');
      
      return err(createError(
        classifyHttpError(response.status, errorText),
        `HTTP ${response.status}: ${errorText || response.statusText}`,
        { retryAfterMs }
      ));
    }

    // Parse JSON response
    const contentType = response.headers.get('content-type') || '';
    let data: T;
    
    if (contentType.includes('application/json')) {
      const text = await response.text();
      data = text ? JSON.parse(text) as T : {} as T;
    } else {
      data = await response.text() as unknown as T;
    }

    return ok({
      status: response.status,
      headers: response.headers,
      data,
    });

  } catch (error) {
    clearTimeout(timeoutId);

    if (error instanceof Error) {
      if (error.name === 'AbortError') {
        return err(createError(
          ErrorCode.TIMEOUT,
          `Request timed out after ${timeoutMs}ms`,
          { retryable: true }
        ));
      }

      if (error.message.includes('ECONNRESET') || 
          error.message.includes('ETIMEDOUT') ||
          error.message.includes('ENOTFOUND')) {
        return err(createError(
          ErrorCode.NETWORK_ERROR,
          error.message,
          { retryable: true }
        ));
      }
    }

    return err(createError(
      ErrorCode.UNKNOWN,
      error instanceof Error ? error.message : String(error),
      { retryable: false }
    ));
  }
}

/**
 * Sleep for a specified duration.
 */
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Clears the rate limit state (for testing).
 */
export function clearRateLimitState(): void {
  rateLimitedUntil.clear();
}
