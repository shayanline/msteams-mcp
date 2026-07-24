/**
 * Unit tests for MSAL v4 localStorage decryption.
 */
import { describe, it, expect } from 'vitest';
import { createCipheriv, hkdfSync, randomBytes } from 'node:crypto';
import { TEAMS_CLIENT_ID } from '../constants.js';
import {
  isEncryptedMsalEntry,
  loadMsalEncryptionKey,
  resolveMsalLocalStorage,
} from './msal-decrypt.js';

function base64Url(buf: Buffer): string {
  return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function encryptMsalEntry(
  plaintext: Record<string, unknown>,
  baseKey: Buffer,
  keyId: string,
  storageKey: string,
): { id: string; nonce: string; data: string } {
  const nonce = randomBytes(16);
  const context = storageKey.includes(TEAMS_CLIENT_ID) ? TEAMS_CLIENT_ID : '';
  const derivedKey = Buffer.from(hkdfSync('sha256', baseKey, nonce, Buffer.from(context, 'utf8'), 32));
  const iv = Buffer.alloc(12, 0);
  const cipher = createCipheriv('aes-256-gcm', derivedKey, iv);
  const body = Buffer.concat([
    cipher.update(Buffer.from(JSON.stringify(plaintext), 'utf8')),
    cipher.final(),
  ]);
  const data = Buffer.concat([body, cipher.getAuthTag()]);
  return {
    id: keyId,
    nonce: base64Url(nonce),
    data: base64Url(data),
  };
}

describe('isEncryptedMsalEntry', () => {
  it('detects v4 encrypted shape and rejects plaintext secrets', () => {
    expect(isEncryptedMsalEntry({ id: 'a', nonce: 'b', data: 'c' })).toBe(true);
    expect(isEncryptedMsalEntry({ id: 'a', nonce: 'b', data: 'c', secret: 'eyJ' })).toBe(false);
    expect(isEncryptedMsalEntry({ secret: 'eyJ', target: 'x' })).toBe(false);
  });
});

describe('loadMsalEncryptionKey', () => {
  it('parses the msal.cache.encryption cookie', () => {
    const key = randomBytes(32);
    const id = 'key-1';
    const raw = encodeURIComponent(JSON.stringify({ id, key: base64Url(key) }));
    const loaded = loadMsalEncryptionKey([{ name: 'msal.cache.encryption', value: raw }]);
    expect(loaded?.id).toBe(id);
    expect(loaded?.key.equals(key)).toBe(true);
  });

  it('returns null when cookie is missing or malformed', () => {
    expect(loadMsalEncryptionKey(undefined)).toBeNull();
    expect(loadMsalEncryptionKey([{ name: 'other', value: 'x' }])).toBeNull();
    expect(loadMsalEncryptionKey([{ name: 'msal.cache.encryption', value: 'not-json' }])).toBeNull();
  });
});

describe('resolveMsalLocalStorage', () => {
  it('decrypts encrypted MSAL entries and passes plaintext through', () => {
    const baseKey = randomBytes(32);
    const keyId = 'enc-1';
    const storageKey = `msal.2|acct|login.windows.net|accesstoken|${TEAMS_CLIENT_ID}|tenant|https://substrate.office.com/SubstrateSearch-Internal.ReadWrite|`;
    const plaintext = {
      credentialType: 'AccessToken',
      secret: 'eyJhbGciOiJub25lIn0.eyJzdWIiOiJ0In0.sig',
      target: 'https://substrate.office.com/SubstrateSearch-Internal.ReadWrite',
      clientId: TEAMS_CLIENT_ID,
      realm: 'tenant',
    };
    const encrypted = encryptMsalEntry(plaintext, baseKey, keyId, storageKey);
    const cookies = [{
      name: 'msal.cache.encryption',
      value: encodeURIComponent(JSON.stringify({ id: keyId, key: base64Url(baseKey) })),
    }];

    const resolved = resolveMsalLocalStorage(
      [
        { name: 'tmp.auth.other', value: '{"item":1}' },
        { name: storageKey, value: JSON.stringify(encrypted) },
        {
          name: 'msal.2|plain',
          value: JSON.stringify({ secret: 'eyJplain', target: 'https://graph.microsoft.com/.default' }),
        },
      ],
      cookies,
    );

    expect(resolved).toHaveLength(3);
    expect(resolved[0]).toEqual({ name: 'tmp.auth.other', value: '{"item":1}' });
    const decrypted = JSON.parse(resolved[1].value);
    expect(decrypted.secret).toBe(plaintext.secret);
    expect(decrypted.target).toBe(plaintext.target);
    expect(decrypted.credentialType).toBe('AccessToken');
    expect(JSON.parse(resolved[2].value).secret).toBe('eyJplain');
  });

  it('passes encrypted entries through unchanged when cookie key id does not match', () => {
    const baseKey = randomBytes(32);
    const storageKey = `msal.2|x|accesstoken|${TEAMS_CLIENT_ID}|t|scope|`;
    const encrypted = encryptMsalEntry(
      { secret: 'eyJ', target: 'scope' },
      baseKey,
      'current-id',
      storageKey,
    );
    const item = { name: storageKey, value: JSON.stringify(encrypted) };
    const resolved = resolveMsalLocalStorage(
      [item],
      [{
        name: 'msal.cache.encryption',
        value: encodeURIComponent(JSON.stringify({ id: 'other-id', key: base64Url(baseKey) })),
      }],
    );
    expect(resolved).toEqual([item]);
  });

  it('passes encrypted entries through unchanged when the encryption cookie is missing', () => {
    const baseKey = randomBytes(32);
    const storageKey = `msal.2|x|accesstoken|${TEAMS_CLIENT_ID}|t|scope|`;
    const encrypted = encryptMsalEntry(
      { secret: 'eyJ', target: 'scope' },
      baseKey,
      'current-id',
      storageKey,
    );
    const item = { name: storageKey, value: JSON.stringify(encrypted) };
    expect(resolveMsalLocalStorage([item], undefined)).toEqual([item]);
    expect(resolveMsalLocalStorage([item], [])).toEqual([item]);
  });

  it('passes encrypted entries through unchanged when decryption fails', () => {
    const baseKey = randomBytes(32);
    const keyId = 'enc-1';
    const storageKey = `msal.2|x|accesstoken|${TEAMS_CLIENT_ID}|t|scope|`;
    const encrypted = encryptMsalEntry(
      { secret: 'eyJ', target: 'scope' },
      baseKey,
      keyId,
      storageKey,
    );
    const item = { name: storageKey, value: JSON.stringify(encrypted) };
    // Matching key id but wrong key material -> auth tag verification fails.
    const resolved = resolveMsalLocalStorage(
      [item],
      [{
        name: 'msal.cache.encryption',
        value: encodeURIComponent(JSON.stringify({ id: keyId, key: base64Url(randomBytes(32)) })),
      }],
    );
    expect(resolved).toEqual([item]);
  });
});
