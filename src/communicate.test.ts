// src/communicate.test.ts
import { describe, expect, it, vi } from 'vitest';
import { runCommunicate, type TeamsStreamer } from './communicate.js';
import { resolveOptions } from './options.js';
import type { ConnectorContext, StreamChunk } from '@legion-collective/core';
import type { MessageActivityInput } from '@microsoft/teams.api';

function fakeCtx(chunks: StreamChunk[]) {
  return {
    streamTool: vi.fn(async () => ({
      conversationId: 'conv-1',
      gen: (async function* () {
        for (const chunk of chunks) yield chunk;
      })(),
    })),
    callTool: vi.fn(),
  } as unknown as ConnectorContext;
}

function fakeStream(): TeamsStreamer & { emitted: string[]; updates: string[]; closed: number } {
  return {
    emitted: [],
    updates: [],
    closed: 0,
    emit(activity) { this.emitted.push(String(activity)); },
    update(text) { this.updates.push(text); },
    async close() { this.closed++; return { id: 'm1' }; },
  };
}

const options = (overrides = {}) => resolveOptions({ stateDir: '/tmp/x', ...overrides });

describe('runCommunicate', () => {
  it('streams text deltas and closes the stream', async () => {
    const stream = fakeStream();
    const send = vi.fn(async () => undefined);
    const result = await runCommunicate({
      ctx: fakeCtx([
        { type: 'iteration_start', iteration: 1 },
        { type: 'text_delta', delta: 'Hello ' },
        { type: 'text_delta', delta: 'world' },
        { type: 'stream:done', result: { status: 'success', data: { conversationId: 'conv-1' } } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), stream, send,
    });
    expect(result.conversationId).toBe('conv-1');
    expect(stream.emitted).toEqual(['Hello ', 'world']);
    expect(stream.updates).toContain('Thinking…');
    expect(stream.closed).toBe(1);
  });

  it('buffers and sends once for minimal detail', async () => {
    const stream = fakeStream();
    const send = vi.fn(async () => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'text_delta', delta: 'Hello ' },
        { type: 'text_delta', delta: 'world' },
        { type: 'stream:done', result: { status: 'success', data: {} } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options({ maxDetail: 'minimal' }), stream, send,
    });
    expect(stream.emitted).toEqual([]);
    expect(send).toHaveBeenCalledTimes(1);
  });

  it('renders approval cards for pending_approval', async () => {
    const send = vi.fn(async () => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        {
          type: 'stream:done',
          result: {
            status: 'pending_approval',
            data: {
              approvalRequests: [
                { approvalId: 'ap-1', requesterId: 'assistant', tool: 'communicate', args: {} },
              ],
            },
          },
        },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options(), stream: fakeStream(), send,
    });
    expect(send).toHaveBeenCalledTimes(1);
  });

  it('shows tool status before text', async () => {
    const stream = fakeStream();
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'tool_call_start', index: 0, id: 't1', name: 'search' },
        { type: 'stream:done', result: { status: 'success', data: {} } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options({ toolDisplay: 'status' }), stream, send: async () => undefined,
    });
    expect(stream.updates).toContain('Calling search…');
  });

  it.each(['status', 'cards'] as const)('hides tool calls at minimal detail with %s display', async (toolDisplay) => {
    const stream = fakeStream();
    const send = vi.fn(async () => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'tool_call_start', index: 0, id: 't1', name: 'search' },
        { type: 'stream:done', result: { status: 'success', data: {} } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options({ maxDetail: 'minimal', toolDisplay }), stream, send,
    });
    expect(stream.updates).toEqual([]);
    expect(send).not.toHaveBeenCalled();
  });

  it('sends only one error for a stream error followed by an error result', async () => {
    const send = vi.fn(async () => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'stream:error', error: 'failed' },
        { type: 'stream:done', result: { status: 'error', error: 'failed' } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options(), stream: fakeStream(), send,
    });
    expect(send).toHaveBeenCalledTimes(1);
  });

  it('reports an empty stream error immediately with a fallback', async () => {
    const send = vi.fn(async (_activity: MessageActivityInput) => undefined);
    const ctx = fakeCtx([]);
    vi.mocked(ctx.streamTool).mockResolvedValue({
      conversationId: 'conv-1',
      gen: (async function* (): AsyncGenerator<StreamChunk> {
        yield { type: 'stream:error', error: '' };
        expect(send).toHaveBeenCalledTimes(1);
        expect(send.mock.calls[0][0].text).toContain('unknown error');
      })(),
    });
    const { result } = await runCommunicate({
      ctx, userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options(), stream: fakeStream(), send,
    });
    expect(send).toHaveBeenCalledTimes(1);
    expect(result.status).toBe('error');
  });

  it('sends buffered text when no streamer is supplied', async () => {
    const send = vi.fn(async (_activity: MessageActivityInput) => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'text_delta', delta: 'Hello ' },
        { type: 'text_delta', delta: 'world' },
        { type: 'stream:done', result: { status: 'success' } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), send,
    });
    expect(send).toHaveBeenCalledTimes(1);
    expect(send.mock.calls[0][0].text).toBe('Hello world');
  });

  it('closes immediately on stream done and only once after generator end', async () => {
    const stream = fakeStream();
    const ctx = fakeCtx([]);
    vi.mocked(ctx.streamTool).mockResolvedValue({
      conversationId: 'conv-1',
      gen: (async function* (): AsyncGenerator<StreamChunk> {
        yield { type: 'stream:done', result: { status: 'success' } };
        expect(stream.closed).toBe(1);
      })(),
    });
    await runCommunicate({
      ctx, userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options(), stream, send: async () => undefined,
    });
    expect(stream.closed).toBe(1);
  });
});
