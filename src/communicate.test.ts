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
const aiLabel = { entities: [expect.objectContaining({ additionalType: ['AIGeneratedContent'] })] };

describe('runCommunicate', () => {
  it('replaces native stream text on a snapshot', async () => {
    const stream = { ...fakeStream(), clearText() { this.emitted.length = 0; } };
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'text_delta', delta: 'draft' },
        { type: 'message_snapshot', content: 'final' },
        { type: 'text_delta', delta: '!' },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), stream,
      send: async () => undefined,
    });
    expect(stream.emitted).toEqual(['final', '!']);
  });

  it('stops typing before the final streamed message', async () => {
    const order: string[] = [];
    const stream: TeamsStreamer = {
      emit() {},
      update() {},
      async close() { order.push('close'); return { id: 'm1' }; },
    };
    await runCommunicate({
      ctx: fakeCtx([{ type: 'stream:done', result: { status: 'success' } }]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), stream,
      send: async () => undefined,
      stopTyping: () => order.push('stopTyping'),
    });
    expect(order).toContain('stopTyping');
    expect(order).toContain('close');
    expect(order.indexOf('stopTyping')).toBeLessThan(order.indexOf('close'));
  });

  it('refreshes one card per tool with its own accumulated arguments', async () => {
    let nextId = 0;
    const send = vi.fn(async (activity: MessageActivityInput) => ({ id: activity.id ?? `card-${++nextId}` }));
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'tool_call_start', index: 0, id: 't1', name: 'search' },
        { type: 'tool_call_args_delta', index: 0, delta: '{"q":' },
        { type: 'tool_call_args_delta', index: 0, delta: '"hi"}' },
        { type: 'tool_call_start', index: 1, id: 't2', name: 'get' },
        { type: 'tool_call_args_delta', index: 1, delta: '{"id":2}' },
        { type: 'stream:done', result: { status: 'success' } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options({ toolDisplay: 'cards' }), send,
    });
    const cards = send.mock.calls.map(([activity]) => ({ id: activity.id, card: activity.attachments?.[0] }));
    const card = (title: string, args: string, id?: string) => ({ id, card: expect.objectContaining({
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: expect.objectContaining({ body: [expect.objectContaining({ text: `🛠️ ${title}` }), expect.objectContaining({ text: args })] }),
    }) });
    expect(cards).toEqual([
      card('search', '(args pending)'), card('search', '{"q":"hi"}', 'card-1'),
      card('get', '(args pending)'), card('get', '{"id":2}', 'card-2'),
    ]);
  });

  it('uses the successful result response as labeled text when no deltas arrived', async () => {
    const send = vi.fn(async (_activity: MessageActivityInput) => undefined);
    const { conversationId } = await runCommunicate({
      ctx: fakeCtx([{ type: 'stream:done', result: { status: 'success', data: { response: 'Answer' } } }]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), stream: fakeStream(), send,
    });
    expect(send).toHaveBeenCalledOnce();
    expect(send).toHaveBeenCalledWith(expect.objectContaining({ text: 'Answer', ...aiLabel }));
    expect(conversationId).toBe('conv-1');
  });

  it.each(['initial send', 'edit'])('closes the stream and returns the thread after a tool card %s fails', async (stage) => {
    const stream = fakeStream();
    const send = vi.fn(async (activity: MessageActivityInput) => {
      if (Boolean(activity.id) === (stage === 'edit')) throw new Error('Bearer secret-token');
      return { id: 'card-1' };
    });
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      const { conversationId } = await runCommunicate({
        ctx: fakeCtx([
          { type: 'tool_call_start', index: 0, id: 't1', name: 'search' },
          { type: 'tool_call_args_delta', index: 0, delta: 'query' },
          { type: 'stream:done', result: { status: 'success', data: { conversationId: 'thread-1' } } },
        ]),
        userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options({ toolDisplay: 'cards' }), stream, send,
      });
      expect(stream.closed).toBe(1);
      expect(conversationId).toBe('thread-1');
      expect(log).toHaveBeenCalled();
      expect(JSON.stringify(log.mock.calls)).not.toContain('secret-token');
    } finally { log.mockRestore(); }
  });

  it('limits accumulated tool card arguments to 500 characters', async () => {
    const send = vi.fn(async (_activity: MessageActivityInput) => ({ id: 'card-1' }));
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'tool_call_start', index: 0, id: 't1', name: 'search' },
        { type: 'tool_call_args_delta', index: 0, delta: 'a'.repeat(250) },
        { type: 'tool_call_args_delta', index: 0, delta: 'b'.repeat(251) },
        { type: 'stream:done', result: { status: 'success' } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options({ toolDisplay: 'cards' }), send,
    });
    expect(send.mock.lastCall?.[0].attachments?.[0].content).toMatchObject({
      body: [expect.anything(), { text: 'a'.repeat(250) + 'b'.repeat(250) }],
    });
  });

  it('sends a labeled partial response once before reporting a timeout error', async () => {
    const send = vi.fn(async (_activity: MessageActivityInput) => undefined);
    const { conversationId } = await runCommunicate({
      ctx: fakeCtx([{ type: 'stream:done', result: {
        status: 'error', error: 'timed out', data: { response: 'partial' },
      } }]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), send,
    });
    expect(send.mock.calls).toEqual([
      [expect.objectContaining({ text: 'partial', ...aiLabel })],
      [expect.objectContaining({ text: 'Error: timed out', ...aiLabel })],
    ]);
    expect(conversationId).toBeUndefined();
  });

  it('labels timeout and error text', async () => {
    const send = vi.fn(async (_activity: MessageActivityInput) => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'call_timeout', timeoutMs: 500 },
        { type: 'stream:done', result: { status: 'error', error: 'failed' } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), send,
    });
    expect(send.mock.calls).toEqual([
      [expect.objectContaining({ text: 'Timed out after 500 ms.', ...aiLabel })],
      [expect.objectContaining({ text: 'Error: failed', ...aiLabel })],
    ]);
  });
  it('streams text deltas and closes the stream', async () => {
    const stream = fakeStream();
    const send = vi.fn(async () => undefined);
    const result = await runCommunicate({
      ctx: fakeCtx([
        { type: 'iteration_start', iteration: 1 },
        { type: 'text_delta', delta: 'Hello ' },
        { type: 'text_delta', delta: 'world' },
        { type: 'stream:done', result: { status: 'success', data: { conversationId: 'thread-1' } } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), stream, send,
    });
    expect(result.conversationId).toBe('thread-1');
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
    expect(send).toHaveBeenCalledWith(expect.objectContaining({ text: 'Hello world', ...aiLabel }));
  });

  it('renders approval cards for pending_approval', async () => {
    const send = vi.fn(async () => undefined);
    const stream = fakeStream();
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
      options: options(), stream, send,
    });
    expect(send).toHaveBeenCalledTimes(1);
    expect(stream.updates).toContain('Awaiting approval…');
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
    expect(send.mock.calls[0][0]).toMatchObject(aiLabel);
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

describe('runCommunicate activity ids', () => {
  it('returns the stored Legion message id and the streamed activity id', async () => {
    const stream = fakeStream();
    const ctx = fakeCtx([
      {
        type: 'stream:done',
        result: {
          status: 'success',
          data: { conversationId: 'thread-1', storedMessageId: 'msg-1' },
        },
      },
    ]);
    const result = await runCommunicate({
      ctx,
      userId: 'alice',
      botOwnerId: 'assistant',
      text: 'hi',
      options: options(),
      stream,
      send: async () => undefined,
    });
    expect(result.storedMessageId).toBe('msg-1');
    expect(result.outboundActivityId).toBe('m1');
  });
});
