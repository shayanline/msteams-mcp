/**
 * Decrypt MSAL Browser v4+ encrypted localStorage cache entries.
 *
 * Starting with MSAL browser v4, tokens in localStorage are AES-GCM encrypted
 * (unless the user selected "Keep me signed in"). The base key lives in the
 * session cookie `msal.cache.encryption`. This mirrors @azure/msal-browser's
 * BrowserCrypto.decrypt / LocalStorage.decryptData.
 *
 * Implemented with node:crypto (sync) so existing sync token extractors can
 * decrypt on read without becoming async.
 */

import { createDecipheriv, hkdfSync } from 'node:crypto';
import { TEAMS_CLIENT_ID } from '../constants.js';
import * as log from '../utils/logger.js';

const ENCRYPTION_COOKIE = 'msal.cache.encryption';
const AES_GCM_IV = Buffer.alloc(12, 0);
const AES_GCM_TAG_LENGTH = 16;

export interface PlaywrightCookie {
  name: string;
  value: string;
  domain?: string;
}

interface EncryptedMsalEntry {
  id: string;
  nonce: string;
  data: string;
  lastUpdatedAt?: string;
}

interface MsalEncryptionKey {
  id: string;
  /** Raw key material used as HKDF IKM. */
  key: Buffer;
}

function base64DecToBuf(base64String: string): Buffer {
  let encoded = base64String.replace(/-/g, '+').replace(/_/g, '/');
  switch (encoded.length % 4) {
    case 0:
      break;
    case 2:
      encoded += '==';
      break;
    case 3:
      encoded += '=';
      break;
    default:
      throw new Error('invalid base64');
  }
  return Buffer.from(encoded, 'base64');
}

export function isEncryptedMsalEntry(entry: unknown): entry is EncryptedMsalEntry {
  if (!entry || typeof entry !== 'object') return false;
  const e = entry as Record<string, unknown>;
  return typeof e.id === 'string'
    && typeof e.nonce === 'string'
    && typeof e.data === 'string'
    && !e.secret;
}

function getEncryptionContext(key: string, clientId = TEAMS_CLIENT_ID): string {
  return key.includes(clientId) ? clientId : '';
}

function decryptPayload(
  baseKey: Buffer,
  nonce: string,
  context: string,
  encryptedData: string,
): string {
  const salt = base64DecToBuf(nonce);
  const info = Buffer.from(context, 'utf8');
  const derivedKey = Buffer.from(hkdfSync('sha256', baseKey, salt, info, 32));
  const encodedData = base64DecToBuf(encryptedData);
  if (encodedData.length <= AES_GCM_TAG_LENGTH) {
    throw new Error('encrypted payload too short');
  }
  const ciphertext = encodedData.subarray(0, encodedData.length - AES_GCM_TAG_LENGTH);
  const authTag = encodedData.subarray(encodedData.length - AES_GCM_TAG_LENGTH);
  const decipher = createDecipheriv('aes-256-gcm', derivedKey, AES_GCM_IV);
  decipher.setAuthTag(authTag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString('utf8');
}

/**
 * Parse the msal.cache.encryption cookie into key material + key id.
 */
export function loadMsalEncryptionKey(
  cookies: PlaywrightCookie[] | undefined,
): MsalEncryptionKey | null {
  const raw = cookies?.find(c => c.name === ENCRYPTION_COOKIE)?.value;
  if (!raw) return null;

  try {
    const parsed = JSON.parse(decodeURIComponent(raw)) as { key?: string; id?: string };
    if (!parsed?.key || !parsed?.id) return null;
    return {
      id: parsed.id,
      key: base64DecToBuf(parsed.key),
    };
  } catch (err) {
    log.debug(
      'msal-decrypt',
      `Failed to parse msal.cache.encryption cookie: ${err instanceof Error ? err.message : String(err)}`,
    );
    return null;
  }
}

/**
 * Resolve MSAL localStorage entries, decrypting v4 encrypted payloads when possible.
 * Non-MSAL keys and already-plaintext MSAL entries are passed through unchanged.
 */
export function resolveMsalLocalStorage(
  localStorage: Array<{ name: string; value: string }>,
  cookies?: PlaywrightCookie[],
): Array<{ name: string; value: string }> {
  const encryption = loadMsalEncryptionKey(cookies);
  let decryptedCount = 0;
  let skippedExpired = 0;
  let decryptFailures = 0;
  const resolved: Array<{ name: string; value: string }> = [];

  for (const item of localStorage) {
    const key = item.name;
    if (!key.startsWith('msal.')) {
      resolved.push(item);
      continue;
    }

    let entry: unknown;
    try {
      entry = JSON.parse(item.value);
    } catch {
      resolved.push(item);
      continue;
    }

    if (!isEncryptedMsalEntry(entry)) {
      resolved.push(item);
      continue;
    }

    // When we cannot decrypt, pass the original encrypted entry through
    // untouched. Callers persist the resolved list back to session-state.json,
    // so dropping entries would permanently lose the MSAL cache. Downstream
    // extractors simply ignore entries without a plaintext secret/target.
    if (!encryption) {
      resolved.push(item);
      decryptFailures++;
      continue;
    }
    if (entry.id !== encryption.id) {
      resolved.push(item);
      skippedExpired++;
      continue;
    }

    try {
      const plaintext = decryptPayload(
        encryption.key,
        entry.nonce,
        getEncryptionContext(key),
        entry.data,
      );
      const decrypted = JSON.parse(plaintext) as Record<string, unknown>;
      if (entry.lastUpdatedAt && !decrypted.lastUpdatedAt) {
        decrypted.lastUpdatedAt = entry.lastUpdatedAt;
      }
      resolved.push({ name: key, value: JSON.stringify(decrypted) });
      decryptedCount++;
    } catch (err) {
      resolved.push(item);
      decryptFailures++;
      log.debug(
        'msal-decrypt',
        `MSAL decrypt failed for key ${key.slice(0, 80)}: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }

  if (decryptedCount || skippedExpired || decryptFailures) {
    log.debug(
      'msal-decrypt',
      `localStorage decrypt summary: decrypted=${decryptedCount}, skippedExpired=${skippedExpired}, failures=${decryptFailures}, hasCookie=${!!encryption}`,
    );
  }

  return resolved;
}
