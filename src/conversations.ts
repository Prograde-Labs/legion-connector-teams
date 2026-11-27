// src/conversations.ts
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import type { ConnectorContext } from '@legion-collective/core';
import { formatToolResult } from './render.js';
import type { TeamsState } from './state.js';

export interface ConversationEntry {
  id: string;
  title: string;
  messageCount: number;
  updatedAt: string;
  status: 'active' | 'archived';
}

interface ConversationMetaLike {
  id?: unknown;
  resolvedTitle?: unknown;
  sharedTitle?: unknown;
  title?: unknown;
  messageCount?: unknown;
  updatedAt?: unknown;
  status?: unknown;
}

export class ConversationToolError extends Error {
  readonly result: { status: string; error?: string; message?: string };

  constructor(result: { status: string; error?: string; message?: string }) {
    super(result.error ?? result.message ?? result.status);
    this.name = 'ConversationToolError';
    this.result = result;
  }
}

function text(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() ? value.trim() : undefined;
}

export async function listConversations(
  ctx: ConnectorContext,
  participantId: string,
): Promise<ConversationEntry[]> {
  const { result } = await ctx.callTool(participantId, 'list_conversations', {
    participantId,
    status: 'active',
  });
  if (result.status !== 'success') throw new ConversationToolError(result);

  const raw = (result.data as { conversations?: unknown } | undefined)?.conversations;
  const list = Array.isArray(raw) ? raw : [];
  return list
    .filter((entry): entry is ConversationMetaLike => typeof entry === 'object' && entry !== null)
    .map((meta) => {
      const id = text(meta.id);
      if (!id) return undefined;
      return {
        id,
        title: text(meta.resolvedTitle) ?? text(meta.sharedTitle) ?? text(meta.title) ?? id,
        messageCount: typeof meta.messageCount === 'number' ? meta.messageCount : 0,
        updatedAt: text(meta.updatedAt) ?? '',
        status: meta.status === 'archived' ? 'archived' : 'active',
      } satisfies ConversationEntry;
    })
    .filter((entry): entry is ConversationEntry => entry !== undefined)
    .sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : a.updatedAt > b.updatedAt ? -1 : 0));
}

export function bindConversation(
  state: TeamsState,
  key: string,
  conversationId: string,
): void {
  const previous = state.legionConversations[key];
  if (previous === conversationId) return;
  if (previous && state.legionConversationTargets[previous] === key) {
    delete state.legionConversationTargets[previous];
  }
  state.legionConversations[key] = conversationId;
  state.legionConversationTargets[conversationId] = key;
}

export function conversationErrorText(error: unknown): string {
  if (error instanceof ConversationToolError) {
    if (error.result.status === 'error') return error.result.error ?? 'unknown error';
    const formatted = formatToolResult(error.result);
    if (formatted) return formatted;
  }
  return error instanceof Error ? error.message : String(error);
}

export const CONVERSATION_PAGE_SIZE = 5;

const MAX_TITLE = 40;

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, max - 1)}…`;
}

function updatedDate(iso: string): string {
  const date = iso.slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : 'unknown';
}

function currentLabel(entries: ConversationEntry[], currentId: string | undefined): string {
  if (!currentId) return 'Current: (none)';
  const current = entries.find((entry) => entry.id === currentId);
  return current
    ? `Current: ${truncate(current.title, MAX_TITLE)}`
    : `Current: ${truncate(currentId, MAX_TITLE)} (not in the active list)`;
}

export function conversationCard(
  entries: ConversationEntry[],
  currentId: string | undefined,
  offset: number,
  key: string,
): IAdaptiveCard {
  const heading = { type: 'TextBlock', text: '🔀 Switch conversation', weight: 'Bolder', wrap: true };
  if (entries.length === 0) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        heading,
        { type: 'TextBlock', text: 'No conversations yet. Send a message to start one.', wrap: true },
      ],
    } as IAdaptiveCard;
  }

  const start = Math.max(0, Math.min(Math.floor(offset), entries.length - 1));
  const page = entries.slice(start, start + CONVERSATION_PAGE_SIZE);
  const body: unknown[] = [
    heading,
    { type: 'TextBlock', text: currentLabel(entries, currentId), wrap: true },
    {
      type: 'TextBlock',
      text: `Showing ${start + 1}–${start + page.length} of ${entries.length}`,
      wrap: true,
    },
    {
      type: 'FactSet',
      facts: page.map((entry) => ({
        title: truncate(entry.title, MAX_TITLE),
        value: `${entry.messageCount} messages · updated ${updatedDate(entry.updatedAt)}` +
          (entry.id === currentId ? ' · current' : ''),
      })),
    },
  ];
  const actions: unknown[] = page.map((entry) => ({
    type: 'Action.Execute',
    title: truncate(entry.title, MAX_TITLE),
    data: { action: 'switch', conversationId: entry.id, key },
  }));
  if (start + CONVERSATION_PAGE_SIZE < entries.length) {
    actions.push({
      type: 'Action.Execute',
      title: 'More…',
      data: { action: 'switch', offset: start + CONVERSATION_PAGE_SIZE, key },
    });
  }
  return { type: 'AdaptiveCard', version: '1.4', body, actions } as IAdaptiveCard;
}

export function switchedCard(
  entry: ConversationEntry,
  alreadyCurrent = false,
): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      {
        type: 'TextBlock',
        text: alreadyCurrent ? '✅ Already on this conversation' : '✅ Switched conversation',
        weight: 'Bolder',
        wrap: true,
      },
      {
        type: 'FactSet',
        facts: [
          { title: 'Title', value: truncate(entry.title, MAX_TITLE) },
          { title: 'Messages', value: String(entry.messageCount) },
          { title: 'Updated', value: updatedDate(entry.updatedAt) },
        ],
      },
      {
        type: 'TextBlock',
        text: 'Future messages in this chat go to this conversation.',
        wrap: true,
      },
    ],
  } as IAdaptiveCard;
}

export function conversationErrorCard(message: string): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [{ type: 'TextBlock', text: `⚠️ ${message}`, wrap: true }],
  } as IAdaptiveCard;
}
