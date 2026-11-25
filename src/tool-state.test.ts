import { describe, expect, it } from 'vitest';
import { emptyState } from './state.js';
import {
  lookupMessageByActivityId,
  lookupMessageByLegionId,
  pruneMessages,
  recordMessage,
} from './tool-state.js';

describe('recordMessage', () => {
  it('records an inbound message and indexes the Legion id', () => {
    const state = emptyState();
    const key = recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-1',
        direction: 'inbound',
        legionMessageId: 'msg-1',
        legionConversationId: 'thread-1',
      },
      10,
    );

    expect(key).toBe('app:act-1');
    expect(state.messages['app:act-1']).toMatchObject({
      appId: 'app',
      conversationId: 'conv-1',
      activityId: 'act-1',
      direction: 'inbound',
      legionMessageId: 'msg-1',
      legionConversationId: 'thread-1',
    });
    expect(state.legionMessages['msg-1']).toBe('app:act-1');
  });

  it('overwrites an existing record with the same key', () => {
    const state = emptyState();
    recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-1',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );
    recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-2',
        activityId: 'act-1',
        direction: 'outbound',
      },
      10,
    );
    expect(state.messages['app:act-1']).toMatchObject({
      conversationId: 'conv-2',
      direction: 'outbound',
    });
    expect(state.messages['app:act-1']?.legionMessageId).toBeUndefined();
  });
});

describe('message lookups', () => {
  it('looks up by Legion id and by activity id', () => {
    const state = emptyState();
    recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-1',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );

    expect(lookupMessageByLegionId(state, 'msg-1')?.activityId).toBe('act-1');
    expect(lookupMessageByActivityId(state, 'act-1', 'app')?.conversationId).toBe('conv-1');
    expect(lookupMessageByActivityId(state, 'act-1', 'other')).toBeUndefined();
    expect(lookupMessageByActivityId(state, 'act-1', 'app', 'conv-2')).toBeUndefined();
  });
});

describe('pruneMessages', () => {
  it('removes the oldest records and their Legion index entries', () => {
    const state = emptyState();
    recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'a',
        direction: 'outbound',
        legionMessageId: 'm-a',
        now: '2026-01-01T00:00:00.000Z',
      },
      2,
    );
    recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'b',
        direction: 'outbound',
        legionMessageId: 'm-b',
        now: '2026-01-02T00:00:00.000Z',
      },
      2,
    );
    recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'c',
        direction: 'outbound',
        legionMessageId: 'm-c',
        now: '2026-01-03T00:00:00.000Z',
      },
      2,
    );

    expect(Object.keys(state.messages).sort()).toEqual(['app:b', 'app:c']);
    expect(state.legionMessages).toEqual({ 'm-b': 'app:b', 'm-c': 'app:c' });
  });

  it('keeps the map unchanged when under the limit', () => {
    const state = emptyState();
    recordMessage(
      state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'a',
        direction: 'outbound',
      },
      5,
    );
    pruneMessages(state, 5);
    expect(Object.keys(state.messages)).toEqual(['app:a']);
  });
});
