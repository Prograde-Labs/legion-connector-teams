// src/communicate.ts
import type { ConnectorContext, StreamChunk, ToolResult } from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { approvalCard, toolCard, type ApprovalRequest } from './render.js';

export interface TeamsStreamer {
  emit(activity: MessageActivityInput | string): void;
  update(text: string): void;
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
  conversationId: string;
}> {
  const { gen, conversationId } = await deps.ctx.streamTool(deps.userId, 'communicate', {
    to: deps.botOwnerId,
    message: deps.text,
    detail: deps.options.maxDetail,
    conversationId: deps.conversationId,
  });

  let buffer = '';
  let final: ToolResult | undefined;
  let sawError: string | undefined;

  for await (const chunk of gen as AsyncGenerator<StreamChunk>) {
    switch (chunk.type) {
      case 'iteration_start':
        if (deps.options.maxDetail !== 'minimal') deps.stream?.update('Thinking…');
        break;
      case 'text_delta':
        buffer += chunk.delta;
        if (deps.options.maxDetail !== 'minimal') deps.stream?.emit(chunk.delta);
        break;
      case 'message_snapshot':
        buffer = chunk.content;
        break;
      case 'tool_call_start':
        if (deps.options.toolDisplay === 'status') {
          deps.stream?.update(`Calling ${chunk.name}…`);
        } else if (deps.options.toolDisplay === 'cards') {
          await deps.send(
            new MessageActivityInput().addCard('adaptive', toolCard(chunk.name, '')),
          );
        }
        break;
      case 'call_timeout':
        await deps.send(new MessageActivityInput(`Timed out after ${chunk.timeoutMs} ms.`));
        break;
      case 'stream:done':
        final = chunk.result;
        break;
      case 'stream:error':
        sawError = chunk.error;
        break;
      default:
        break;
    }
  }

  if (sawError) {
    await deps.send(new MessageActivityInput(`Error: ${sawError}`));
    final ??= { status: 'error', error: sawError };
  }

  if (deps.options.maxDetail === 'minimal') {
    if (buffer) await deps.send(new MessageActivityInput(buffer));
  } else {
    await deps.stream?.close();
  }

  if (final?.status === 'pending_approval') {
    const data = final.data as { approvalRequests?: ApprovalRequest[] } | undefined;
    for (const approval of data?.approvalRequests ?? []) {
      await deps.send(new MessageActivityInput().addCard('adaptive', approvalCard(approval)));
    }
  } else if (final?.status === 'error') {
    await deps.send(new MessageActivityInput(`Error: ${final.error ?? 'unknown error'}`));
  } else if (final?.status === 'rejected') {
    await deps.send(new MessageActivityInput(final.message ?? 'Rejected.'));
  }

  return { result: final ?? { status: 'success' }, conversationId };
}
