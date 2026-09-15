/**
 * Loop Web Service client.
 *
 * This is the high-level service used by Microsoft Loop clients for creating,
 * reading, and modifying pages. Unlike the SharePoint HTML export endpoint, it
 * accepts Markdown-like raw content and materialises the corresponding Fluid
 * operations server-side.
 */

import { randomUUID } from 'node:crypto';
import { getLoopApiToken } from '../auth/index.js';
import { LOOP_ORIGIN, LOOP_USER_AGENT, LOOP_WEB_SERVICE_BASE } from '../constants.js';
import { fetchWithRetry, parseResponse } from '../utils/http.js';
import type { SpoCoordinates } from '../types/loop.js';
import { isSharePointHost } from '../utils/parsers.js';

export type PagePosition = 'first' | 'last' | 'before' | 'after';

export interface RawPageContent {
  type: 'raw';
  value: string;
  insertionOptions?: {
    attribution: 'Copilot';
    isAIGenerated: boolean;
  };
}

/**
 * A page's position within its workspace's page tree: either as the first or
 * last child of a parent (root-level when `parent` is omitted), or immediately
 * before/after a sibling element. Used both to place a page at creation time
 * and, via `moveLoopPage`, to reposition an already-created page.
 */
export type PageTreeLocation =
  | { type: 'first' | 'last'; parent?: string }
  | { type: 'before' | 'after'; sibling: string };

export interface CreatePageRequest {
  title: string;
  content: RawPageContent;
  location: PageTreeLocation;
  shareLinkOptions?: { scope: 'organization' | 'users' | 'default' };
}

/**
 * Reposition an existing page within its workspace's tree (move to a new
 * parent, or before/after a different sibling). Content and title are left
 * untouched — use `modifyLoopPage` for those.
 */
export interface MovePageRequest {
  location: PageTreeLocation;
}

export interface CreatePageResponse {
  page: {
    pageId: string;
    elementId: string;
    link: string;
    shareLink?: string;
  };
}

export interface ModifyPageRequest {
  title?: string;
  content?: { type: 'raw'; value: string };
  location?: {
    type: 'before' | 'after' | 'replace';
    path?: Array<{ type: 'TargetLabel'; value: string }> | 'REPLACE_ALL';
  };
}

export interface ReadPageResponse {
  title: string;
  content: string;
}

export interface LoopWebPageListItem {
  type: 'Loop' | 'Link';
  title: string;
  link?: string;
  elementId: string;
  children?: LoopWebPageListItem[];
  odspMetadata?: {
    siteUrl: string;
    driveId: string;
    itemId: string;
  };
}

export interface ListPagesResponse {
  pages: LoopWebPageListItem[];
}

/** Encode the page locator expected by Loop Web Service. */
export function encodeLoopWebPageId({ host, driveId, itemId }: SpoCoordinates): string {
  if (!isSharePointHost(host)) throw new Error(`Invalid SharePoint host "${host}".`);
  return Buffer.from(`${host},${driveId},${itemId}`, 'utf8').toString('base64');
}

/** Decode and validate a Loop Web Service page id. */
export function decodeLoopWebPageId(pageId: string): SpoCoordinates | null {
  try {
    const decoded = Buffer.from(pageId, 'base64').toString('utf8');
    const [host, driveId, itemId, ...extra] = decoded.split(',');
    if (extra.length > 0 || !isSharePointHost(host) || !driveId || !itemId) return null;
    return { host, driveId, itemId };
  } catch {
    return null;
  }
}

function segment(value: string): string {
  return encodeURIComponent(value);
}

async function loopWebRequest<T>(method: string, path: string, body?: unknown, maxAttempts = 1): Promise<T> {
  const token = await getLoopApiToken();
  if (!token) {
    throw new Error(
      'Loop page-write token unavailable. Run loop_login, or sign in with Azure CLI using ' +
      '`az login --scope https://api.loop.cloud.microsoft/.default --allow-no-subscriptions`.',
    );
  }

  const res = await fetchWithRetry(`${LOOP_WEB_SERVICE_BASE}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
      Origin: LOOP_ORIGIN,
      Referer: `${LOOP_ORIGIN}/`,
      'User-Agent': LOOP_USER_AGENT,
      'MS-CV': randomUUID(),
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  }, maxAttempts);

  try {
    return await parseResponse<T>(res);
  } catch (err) {
    throw new Error(`Loop Web Service request failed: ${err instanceof Error ? err.message : String(err)}`);
  }
}

export function createLoopPage(workspaceId: string, request: CreatePageRequest): Promise<CreatePageResponse> {
  return loopWebRequest<CreatePageResponse>('POST', `/workspaces/${segment(workspaceId)}/pages`, request);
}

export function modifyLoopPage(pageId: string, request: ModifyPageRequest): Promise<void> {
  return loopWebRequest<void>('PATCH', `/pages/${segment(pageId)}`, request);
}

/**
 * Move an existing page to a new position in its workspace's tree (reparent,
 * or reorder relative to a sibling). Uses the same PATCH endpoint as
 * `modifyLoopPage`, but with the tree-placement `location` shape the service
 * accepts at creation time, rather than the content-insertion `location`
 * shape `modifyLoopPage` uses.
 */
export function moveLoopPage(pageId: string, request: MovePageRequest): Promise<void> {
  return loopWebRequest<void>('PATCH', `/pages/${segment(pageId)}`, request);
}

export function readLoopPage(pageId: string): Promise<ReadPageResponse> {
  return loopWebRequest<ReadPageResponse>('POST', `/pages/${segment(pageId)}`, {
    options: { type: 'markdown' },
  }, 3);
}

export function listLoopPages(workspaceId: string): Promise<ListPagesResponse> {
  return loopWebRequest<ListPagesResponse>('POST', `/workspaces/${segment(workspaceId)}/pages/list`, {}, 3);
}
