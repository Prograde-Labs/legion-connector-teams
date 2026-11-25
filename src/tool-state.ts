import type { MessageRecord, TeamsState } from './state.js';

export interface RecordMessageInput {
  appId: string;
  conversationId: string;
  activityId: string;
  direction: 'inbound' | 'outbound';
  legionMessageId?: string;
  legionConversationId?: string;
  now?: string;
}

export function messageKey(appId: string, activityId: string): string {
  return `${appId}:${activityId}`;
}

export function recordMessage(
  state: TeamsState,
  input: RecordMessageInput,
  maxRecords: number,
): string {
  const key = messageKey(input.appId, input.activityId);
  state.messages[key] = {
    appId: input.appId,
    conversationId: input.conversationId,
    activityId: input.activityId,
    direction: input.direction,
    ...(input.legionMessageId === undefined ? {} : { legionMessageId: input.legionMessageId }),
    ...(input.legionConversationId === undefined
      ? {}
      : { legionConversationId: input.legionConversationId }),
    createdAt: input.now ?? new Date().toISOString(),
  };
  if (input.legionMessageId !== undefined) {
    state.legionMessages[input.legionMessageId] = key;
  }
  pruneMessages(state, maxRecords);
  return key;
}

export function lookupMessageByLegionId(
  state: TeamsState,
  legionMessageId: string,
): MessageRecord | undefined {
  const key = state.legionMessages[legionMessageId];
  return key === undefined ? undefined : state.messages[key];
}

export function lookupMessageByActivityId(
  state: TeamsState,
  activityId: string,
  appId?: string,
  conversationId?: string,
): MessageRecord | undefined {
  for (const record of Object.values(state.messages)) {
    if (record.activityId !== activityId) continue;
    if (appId !== undefined && record.appId !== appId) continue;
    if (conversationId !== undefined && record.conversationId !== conversationId) continue;
    return record;
  }
  return undefined;
}

export function pruneMessages(state: TeamsState, maxRecords: number): void {
  const keys = Object.keys(state.messages);
  if (keys.length <= maxRecords) return;

  const ordered = keys.sort((a, b) => {
    const left = state.messages[a]?.createdAt ?? '';
    const right = state.messages[b]?.createdAt ?? '';
    return left === right ? a.localeCompare(b) : left.localeCompare(right);
  });

  for (const key of ordered.slice(0, keys.length - maxRecords)) {
    const record = state.messages[key];
    delete state.messages[key];
    if (
      record?.legionMessageId !== undefined &&
      state.legionMessages[record.legionMessageId] === key
    ) {
      delete state.legionMessages[record.legionMessageId];
    }
  }
}
