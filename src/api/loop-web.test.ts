import { beforeEach, describe, expect, it, vi } from 'vitest';
import { createLoopPage, decodeLoopWebPageId, encodeLoopWebPageId, listLoopPages, modifyLoopPage, readLoopPage } from './loop-web.js';
import { LOOP_CLIENT_SCENARIO } from '../constants.js';

vi.mock('../auth/index.js', () => ({
  getLoopApiToken: vi.fn(async () => 'loop-token'),
}));

describe('encodeLoopWebPageId', () => {
  it('encodes host, drive, and item coordinates as base64', () => {
    const encoded = encodeLoopWebPageId({ host: 'contoso.sharepoint.com', driveId: 'drive-1', itemId: 'item-2' });
    expect(Buffer.from(encoded, 'base64').toString('utf8')).toBe('contoso.sharepoint.com,drive-1,item-2');
  });

  it('rejects page ids containing non-SharePoint hosts', () => {
    expect(() => encodeLoopWebPageId({
      host: 'attacker.example.com',
      driveId: 'drive-1',
      itemId: 'item-2',
    })).toThrow('Invalid SharePoint host');
    expect(decodeLoopWebPageId(
      Buffer.from('attacker.example.com,drive-1,item-2').toString('base64'),
    )).toBeNull();
  });

  it('round-trips valid ids and rejects malformed ids', () => {
    const coordinates = { host: 'contoso.sharepoint.com', driveId: 'drive-1', itemId: 'item-2' };
    expect(decodeLoopWebPageId(encodeLoopWebPageId(coordinates))).toEqual(coordinates);
    expect(decodeLoopWebPageId(Buffer.from('not,enough').toString('base64'))).toBeNull();
  });
});

describe('Loop Web Service requests', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it('creates a page with authenticated JSON and an encoded workspace id', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(JSON.stringify({
      page: { pageId: 'page-1', elementId: 'element-1', link: 'https://loop.cloud.microsoft/p/page-1' },
    }), { status: 200, headers: { 'content-type': 'application/json' } }));

    await createLoopPage('pod/id=', {
      title: 'Notes',
      content: { type: 'raw', value: '# Hello' },
      location: { type: 'last' },
    });

    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toContain('/workspaces/pod%2Fid%3D/pages');
    expect(init?.method).toBe('POST');
    expect((init?.headers as Record<string, string>).Authorization).toBe('Bearer loop-token');
    expect(JSON.parse(init?.body as string)).toMatchObject({ title: 'Notes', location: { type: 'last' } });
  });

  it('sends replace-section updates', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(null, { status: 204 }));
    await modifyLoopPage('page/id=', {
      content: { type: 'raw', value: 'Updated' },
      location: { type: 'replace', path: [{ type: 'TargetLabel', value: 'Summary' }] },
    });
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toContain('/pages/page%2Fid%3D');
    expect(init?.method).toBe('PATCH');
  });

  it('requests Markdown when reading a page', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(
      JSON.stringify({ title: 'Notes', content: '# Notes' }),
      { status: 200, headers: { 'content-type': 'application/json' } },
    ));
    await expect(readLoopPage('page')).resolves.toEqual({ title: 'Notes', content: '# Notes' });
    expect(JSON.parse(fetchMock.mock.calls[0][1]?.body as string)).toEqual({ options: { type: 'markdown' } });
  });

  it('lists pages in a workspace', async () => {
    const fetchMock = vi.spyOn(globalThis, 'fetch').mockResolvedValue(new Response(
      JSON.stringify({ pages: [] }),
      { status: 200, headers: { 'content-type': 'application/json' } },
    ));
    await expect(listLoopPages('pod/id=')).resolves.toEqual({ pages: [] });
    expect(fetchMock.mock.calls[0][0]).toContain('/workspaces/pod%2Fid%3D/pages/list');
    const headers = fetchMock.mock.calls[0][1]?.headers as Record<string, string>;
    expect(headers['X-Client-Scenario']).toBe(LOOP_CLIENT_SCENARIO);
  });
});
