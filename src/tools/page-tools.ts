/**
 * Page MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { discover } from '../api/loop.js';
import { getPageContent } from '../api/pages.js';
import { resolvePageCoordinates } from '../api/pages.js';
import {
  createLoopPage,
  decodeLoopWebPageId,
  encodeLoopWebPageId,
  listLoopPages,
  modifyLoopPage,
  moveLoopPage,
  readLoopPage,
  type ModifyPageRequest,
  type LoopWebPageListItem,
  type PageTreeLocation,
} from '../api/loop-web.js';
import { decodePodId } from '../utils/parsers.js';
import type { LoopPage, LoopWorkspace } from '../types/loop.js';

function summarisePage(p: LoopPage) {
  return {
    id: p.id,
    title: p.title ?? '(untitled)',
    type: p.type,
    workspaceId: p.workspace_id,
    isDeleted: p.is_deleted ?? false,
  };
}

/** Fallback SharePoint coordinates for a page, from its workspace pod_id. */
function workspaceFallback(workspaces: LoopWorkspace[], workspaceId: string | undefined) {
  if (!workspaceId) return null;
  const ws = workspaces.find(w => w.id === workspaceId);
  return decodePodId(ws?.mfs_info?.pod_id);
}

function toolResult(value: unknown) {
  return { content: [{ type: 'text' as const, text: JSON.stringify(value, null, 2) }] };
}

function errorResult(action: string, err: unknown) {
  return toolResult({
    success: false,
    message: `${action}: ${err instanceof Error ? err.message : String(err)}`,
  });
}

function findWorkspace(workspaces: LoopWorkspace[], workspaceId: string): LoopWorkspace | undefined {
  return workspaces.find(w => w.id === workspaceId || w.mfs_info?.pod_id === workspaceId);
}

type TreeLocationResult =
  | { ok: true; location: PageTreeLocation }
  | { ok: false; message: string };

/**
 * Build a page-tree location from position + parent/sibling element ids.
 * Shared by loop_create_page (placement at creation) and loop_move_page
 * (repositioning an existing page) since both accept the same shape.
 */
function buildTreeLocation(
  position: 'first' | 'last' | 'before' | 'after',
  parentElementId: string | undefined,
  siblingElementId: string | undefined,
): TreeLocationResult {
  if (position === 'before' || position === 'after') {
    if (!siblingElementId) {
      return { ok: false, message: `sibling_element_id is required with position ${position}.` };
    }
    return { ok: true, location: { type: position, sibling: siblingElementId } };
  }
  return {
    ok: true,
    location: parentElementId ? { type: position, parent: parentElementId } : { type: position },
  };
}

function flattenLoopWebPages(
  pages: LoopWebPageListItem[],
  workspaceId: string,
  parentElementId?: string,
  depth = 0,
): Array<Record<string, unknown>> {
  return pages.flatMap(page => {
    const metadata = page.odspMetadata;
    const current = metadata && page.type === 'Loop'
      ? [{
          id: encodeLoopWebPageId({
            host: new URL(metadata.siteUrl).hostname,
            driveId: metadata.driveId,
            itemId: metadata.itemId,
          }),
          elementId: page.elementId,
          title: page.title,
          type: page.type,
          workspaceId,
          parentElementId,
          depth,
          link: page.link,
        }]
      : [];
    return [
      ...current,
      ...flattenLoopWebPages(page.children ?? [], workspaceId, page.elementId, depth + 1),
    ];
  });
}

export function registerPageTools(server: McpServer): void {
  // ── loop_list_pages ──────────────────────────────────────────────────────
  server.tool(
    'loop_list_pages',
    'List the pages in a Loop workspace. Pass the workspace id from loop_list_workspaces.',
    {
      workspace_id: z.string().min(1).describe('The workspace id to list pages for.'),
      include_deleted: z.boolean().optional().describe('Include deleted pages. Default: false.'),
    },
    async ({ workspace_id, include_deleted }) => {
      try {
        const { pages, workspaces } = await discover();
        const workspace = findWorkspace(workspaces, workspace_id);

        if (workspace && !include_deleted) {
          try {
            const result = await listLoopPages(workspace.mfs_info?.pod_id ?? workspace.id);
            const listed = flattenLoopWebPages(result.pages, workspace.id);
            return toolResult({ count: listed.length, source: 'loop-web-service', pages: listed });
          } catch {
            // Fall through to Substrate metadata when Loop Web Service is unavailable.
          }
        }

        const filtered = pages.filter(
          p => p.workspace_id === workspace_id && (include_deleted || !p.is_deleted),
        );
        return toolResult({ count: filtered.length, source: 'substrate', pages: filtered.map(summarisePage) });
      } catch (err) {
        return errorResult('Page listing failed', err);
      }
    },
  );

  // ── loop_get_page ────────────────────────────────────────────────────────
  server.tool(
    'loop_get_page',
    'Read the content of a Loop page as Markdown. Pass the page id from loop_list_pages. Loop pages are rich Fluid documents exported to HTML, so some interactive components may render approximately.',
    {
      page_id: z.string().min(1).describe('The page id to read.'),
      format: z.enum(['markdown', 'html']).optional().describe('Output format. Default: markdown.'),
    },
    async ({ page_id, format }) => {
      const directCoordinates = decodeLoopWebPageId(page_id);
      if (directCoordinates) {
        if (format === 'html') {
          return toolResult({
            success: false,
            message: 'HTML output is unavailable for a Loop Web Service page id; use format=markdown.',
          });
        }
        try {
          const page = await readLoopPage(page_id);
          return toolResult({ success: true, id: page_id, title: page.title, content: page.content });
        } catch (err) {
          return errorResult('Page read failed', err);
        }
      }

      const { pages, workspaces } = await discover();
      const page = pages.find(p => p.id === page_id);
      if (!page) {
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({ success: false, message: `No page found with id ${page_id}.` }, null, 2),
          }],
        };
      }

      const fallback = workspaceFallback(workspaces, page.workspace_id);
      const content = await getPageContent(page, fallback);

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            id: page.id,
            title: page.title ?? '(untitled)',
            workspaceId: page.workspace_id,
            content: format === 'html' ? content.html : content.markdown,
          }, null, 2),
        }],
      };
    },
  );

  // ── loop_create_page ────────────────────────────────────────────────────
  server.tool(
    'loop_create_page',
    'Create a Microsoft Loop page from Markdown content. The page is created through the same Loop Web Service used by Microsoft clients.',
    {
      workspace_id: z.string().min(1).describe('Workspace id from loop_list_workspaces.'),
      title: z.string().min(1).describe('Page title.'),
      content: z.string().optional().describe('Initial page content in Markdown. Default: empty page.'),
      position: z.enum(['first', 'last', 'before', 'after']).optional().describe('Position in the workspace. Default: last.'),
      parent_element_id: z.string().min(1).optional().describe('Create as a child of this page element id; valid with first/last.'),
      sibling_element_id: z.string().min(1).optional().describe('Sibling element id; required with before/after.'),
      share_scope: z.enum(['organization', 'users', 'default']).optional().describe('Optionally create a sharing link with this scope.'),
    },
    async ({ workspace_id, title, content, position, parent_element_id, sibling_element_id, share_scope }) => {
      try {
        const { workspaces } = await discover();
        const workspace = findWorkspace(workspaces, workspace_id);
        if (!workspace) return toolResult({ success: false, message: `No workspace found with id ${workspace_id}.` });

        const apiWorkspaceId = workspace.mfs_info?.pod_id ?? workspace.id;
        const selectedPosition = position ?? 'last';
        const locationResult = buildTreeLocation(selectedPosition, parent_element_id, sibling_element_id);
        if (!locationResult.ok) return toolResult({ success: false, message: locationResult.message });

        const created = await createLoopPage(apiWorkspaceId, {
          title,
          content: { type: 'raw', value: content ?? '' },
          location: locationResult.location,
          ...(share_scope ? { shareLinkOptions: { scope: share_scope } } : {}),
        });
        return toolResult({ success: true, message: `Created Loop page "${title}".`, page: created.page });
      } catch (err) {
        return errorResult('Page creation failed', err);
      }
    },
  );

  // ── loop_move_page ──────────────────────────────────────────────────────
  server.tool(
    'loop_move_page',
    'Move an existing Loop page to a new position in its workspace tree — reparent it under a different page, or reorder it relative to a sibling. Does not change the page title or content.',
    {
      page_id: z.string().min(1).describe('Page id from loop_list_pages or loop_create_page.'),
      position: z.enum(['first', 'last', 'before', 'after']).describe('New position. first/last are relative to parent_element_id (or the workspace root if omitted); before/after require sibling_element_id.'),
      parent_element_id: z.string().min(1).optional().describe('Move as a child of this page element id; valid with first/last. Omit to place at the workspace root.'),
      sibling_element_id: z.string().min(1).optional().describe('Sibling element id to move before/after; required with before/after.'),
    },
    async ({ page_id, position, parent_element_id, sibling_element_id }) => {
      try {
        const locationResult = buildTreeLocation(position, parent_element_id, sibling_element_id);
        if (!locationResult.ok) return toolResult({ success: false, message: locationResult.message });

        let coordinates = decodeLoopWebPageId(page_id);
        if (!coordinates) {
          const { pages, workspaces } = await discover();
          const page = pages.find(p => p.id === page_id);
          if (!page) return toolResult({ success: false, message: `No page found or invalid Loop page id: ${page_id}.` });
          const resolved = resolvePageCoordinates(page, workspaceFallback(workspaces, page.workspace_id));
          if (!resolved) return toolResult({ success: false, message: `Could not resolve SharePoint coordinates for page ${page_id}.` });
          coordinates = resolved;
        }

        await moveLoopPage(encodeLoopWebPageId(coordinates), { location: locationResult.location });
        return toolResult({ success: true, message: `Moved Loop page ${page_id}.`, id: page_id });
      } catch (err) {
        return errorResult('Page move failed', err);
      }
    },
  );

  // ── loop_update_page ────────────────────────────────────────────────────
  server.tool(
    'loop_update_page',
    'Update a Loop page. Appends by default; can prepend, replace a named heading section, replace the complete body, or change the title. Full replacement requires confirm_replace_all=true.',
    {
      page_id: z.string().min(1).describe('Page id from loop_list_pages or loop_create_page.'),
      title: z.string().min(1).optional().describe('New page title.'),
      content: z.string().min(1).optional().describe('Markdown content to insert or replace.'),
      operation: z.enum(['append', 'prepend', 'replace_section', 'replace_all']).optional().describe('Content operation. Default: append.'),
      section_heading: z.string().min(1).optional().describe('Heading name required by replace_section.'),
      confirm_replace_all: z.boolean().optional().describe('Must be true for replace_all because it removes the current page body.'),
    },
    async ({ page_id, title, content, operation, section_heading, confirm_replace_all }) => {
      try {
        if (!title && !content) return toolResult({ success: false, message: 'Provide title, content, or both.' });

        let page: LoopPage | undefined;
        let coordinates = decodeLoopWebPageId(page_id);
        if (!coordinates) {
          const { pages, workspaces } = await discover();
          page = pages.find(p => p.id === page_id);
          coordinates = page
            ? resolvePageCoordinates(page, workspaceFallback(workspaces, page.workspace_id))
            : null;
        }
        if (!coordinates) return toolResult({ success: false, message: `No page found or invalid Loop page id: ${page_id}.` });
        const apiPageId = page ? encodeLoopWebPageId(coordinates) : page_id;

        const request: ModifyPageRequest = {};
        if (title) request.title = title;
        if (content) {
          request.content = { type: 'raw', value: content };
          const selectedOperation = operation ?? 'append';
          if (selectedOperation === 'replace_section') {
            if (!section_heading) {
              return toolResult({ success: false, message: 'section_heading is required for replace_section.' });
            }
            request.location = { type: 'replace', path: [{ type: 'TargetLabel', value: section_heading }] };
          } else if (selectedOperation === 'replace_all') {
            if (confirm_replace_all !== true) {
              return toolResult({ success: false, message: 'Set confirm_replace_all=true to replace the complete page body.' });
            }
            request.location = { type: 'replace', path: 'REPLACE_ALL' };
          } else {
            request.location = { type: selectedOperation === 'prepend' ? 'before' : 'after' };
          }
        }

        await modifyLoopPage(apiPageId, request);
        try {
          const updated = await readLoopPage(apiPageId);
          return toolResult({
            success: true,
            message: `Updated Loop page "${updated.title}".`,
            id: page?.id ?? page_id,
            title: updated.title,
            content: updated.content,
          });
        } catch (verificationError) {
          return toolResult({
            success: true,
            message: 'The update was accepted, but the read-back verification failed. Do not repeat an append automatically.',
            id: page?.id ?? page_id,
            verificationError: verificationError instanceof Error ? verificationError.message : String(verificationError),
          });
        }
      } catch (err) {
        return errorResult('Page update failed', err);
      }
    },
  );
}
