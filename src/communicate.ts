// src/communicate.ts
import type { ConnectorContext, StreamChunk, ToolResult } from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { approvalCard, sendText, toolCard, type ApprovalRequest } from './render.js';

const MAX_TOOL_CARD_ARGS = 500;

export interface TeamsStreamer {
  emit(activity: MessageActivityInput | string): void;
  update(text: string): void;
  clearText?(): void;
  close(): Promise<{ id?: string } | undefined>;
}

export interface CommunicateDeps {
  ctx: ConnectorContext;
  userId: string;
  botOwnerId: string;
  text: string;
  conversationId?: string;
  options: ResolvedOptions;
  stream?: TeamsStreamer;
  send: (activity: MessageActivityInput) => Promise<unknown>;
}

export async function runCommunicate(deps: CommunicateDeps): Promise<{
  result: ToolResult;
  conversationId: string | undefined;
}> {
  const { gen, conversationId } = await deps.ctx.streamTool(deps.userId, 'communicate', {
    to: deps.botOwnerId,
    message: deps.text,
    detail: deps.options.maxDetail,
    conversationId: deps.conversationId,
  });

  let buffer = '';
  let final: ToolResult | undefined;
  let sawError = false;
  let errorText = '';
  let errorSent = false;
  let closed = false;
  const reply = (text: string) => sendText(deps.send, deps.options.aiLabel, text);
  const tools = new Map<number, { name: string; args: string; displayedArgs?: string; activityId?: string }>();

  async function sendCard(card: MessageActivityInput): Promise<{ id?: string } | undefined> {
    try {
      return await deps.send(card) as { id?: string } | undefined;
    } catch {
      // SDK errors may include credentials or the card payload. Do not log them.
      console.error(card.id ? 'Teams card update failed' : 'Teams card send failed');
      return undefined;
    }
  }

  async function refreshToolCards(): Promise<void> {
    for (const tool of tools.values()) {
      if (tool.displayedArgs === tool.args) continue;
      const card = new MessageActivityInput().addCard('adaptive', toolCard(tool.name, tool.args));
      // Sending with a known ID updates the existing Teams activity.
      if (tool.activityId) card.id = tool.activityId;
      const sent = await sendCard(card);
      tool.activityId = sent?.id ?? tool.activityId;
      tool.displayedArgs = tool.args;
    }
  }

  for await (const chunk of gen as AsyncGenerator<StreamChunk>) {
    switch (chunk.type) {
      case 'iteration_start':
        await refreshToolCards();
        tools.clear();
        if (deps.options.maxDetail !== 'minimal') deps.stream?.update('Thinking…');
        break;
      case 'text_delta':
        buffer += chunk.delta;
        if (deps.options.maxDetail !== 'minimal') deps.stream?.emit(chunk.delta);
        break;
      case 'message_snapshot':
        buffer = chunk.content;
        if (deps.options.maxDetail !== 'minimal' && deps.stream) {
          deps.stream.clearText?.();
          deps.stream.emit(buffer);
        }
        break;
      case 'tool_call_start':
        if (deps.options.maxDetail === 'minimal') break;
        if (deps.options.toolDisplay === 'status') {
          deps.stream?.update(`Calling ${chunk.name}…`);
        } else if (deps.options.toolDisplay === 'cards') {
          await refreshToolCards();
          const sent = await sendCard(
            new MessageActivityInput().addCard('adaptive', toolCard(chunk.name, '(args pending)')),
          );
          tools.set(chunk.index, { name: chunk.name, args: '', activityId: sent?.id });
        }
        break;
      case 'tool_call_args_delta': {
        const tool = tools.get(chunk.index);
        if (tool) tool.args += chunk.delta.slice(0, MAX_TOOL_CARD_ARGS - tool.args.length);
        break;
      }
      case 'call_timeout':
        await reply(`Timed out after ${chunk.timeoutMs} ms.`);
        break;
      case 'stream:done':
        final = chunk.result;
        await refreshToolCards();
        if (deps.options.maxDetail !== 'minimal' && !closed && deps.stream) {
          // Replace a stale "Calling <tool>…" status when the turn paused for
          // approval; the resumed reply arrives as a separate message.
          if (final.status === 'pending_approval') deps.stream.update('Awaiting approval…');
          closed = true;
          await deps.stream.close();
        }
        break;
      case 'stream:error':
        sawError = true;
        errorText = chunk.error || 'unknown error';
        if (!errorSent) {
          errorSent = true;
          await reply(`Error: ${errorText}`);
        }
        break;
      default:
        break;
    }
  }
  await refreshToolCards();

  if (sawError) {
    final ??= { status: 'error', error: errorText };
  }

  if (deps.options.maxDetail === 'minimal' || !deps.stream) {
    if (buffer) await reply(buffer);
  } else if (!closed) {
    closed = true;
    await deps.stream.close();
  }

  const data = final?.data as {
    conversationId?: unknown;
    response?: unknown;
    approvalRequests?: ApprovalRequest[];
  } | undefined;
  if (!buffer && typeof data?.response === 'string') {
    await reply(data.response);
  }

  if (final?.status === 'pending_approval') {
    for (const approval of data?.approvalRequests ?? []) {
      await sendCard(new MessageActivityInput().addCard('adaptive', approvalCard(approval)));
    }
  } else if (final?.status === 'error' && !errorSent) {
    await reply(`Error: ${final.error ?? 'unknown error'}`);
  } else if (final?.status === 'rejected') {
    await reply(final.message ?? 'Rejected.');
  }

  return {
    result: final ?? { status: 'success' },
    conversationId: typeof data?.conversationId === 'string'
      ? data.conversationId
      : final?.status === 'error' ? undefined : conversationId,
  };
}
