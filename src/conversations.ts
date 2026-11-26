// src/conversations.ts
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
