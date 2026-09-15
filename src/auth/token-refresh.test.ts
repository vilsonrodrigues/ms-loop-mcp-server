import { describe, expect, it } from 'vitest';
import { inferSharePointResource } from './token-refresh.js';

const podId = (value: string) => Buffer.from(value, 'utf8').toString('base64');

describe('inferSharePointResource', () => {
  it('derives the resource from a workspace pod id', () => {
    expect(inferSharePointResource([{
      id: 'workspace-1',
      mfs_info: { pod_id: podId('prefix|contoso.sharepoint.com|b!drive|01ITEM') },
    }])).toBe('https://contoso.sharepoint.com');
  });

  it('ignores malformed and non-SharePoint pod ids', () => {
    expect(inferSharePointResource([
      { id: 'bad', mfs_info: { pod_id: podId('prefix|attacker.example.com|drive|item') } },
      { id: 'missing' },
    ])).toBeNull();
  });
});
