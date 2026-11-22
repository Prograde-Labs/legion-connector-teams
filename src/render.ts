import type { IAdaptiveCard } from '@microsoft/teams.cards';
import { MessageActivityInput } from '@microsoft/teams.api';

export async function sendText(
  send: (activity: MessageActivityInput) => Promise<unknown>,
  aiLabel: boolean,
  text: string,
): Promise<void> {
  const activity = new MessageActivityInput(text);
  if (aiLabel) activity.addAiGenerated();
  await send(activity);
}

export interface ApprovalRequest {
  approvalId: string;
  conversationId?: string;
  requesterId: string;
  tool: string;
  args: unknown;
  createdAt?: string;
}

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, max - 1)}…`;
}

function serializeArgs(args: unknown): string {
  try {
    return JSON.stringify(args) ?? '<unserializable args>';
  } catch {
    return '<unserializable args>';
  }
}

export function chunkMarkdown(text: string, maxChars: number): string[] {
  const limit = Number.isFinite(maxChars) ? Math.max(1, Math.floor(maxChars)) : 1;
  if (text.length <= limit) return [text];
  const chunks: string[] = [];
  let remaining = text;
  while (remaining.length > limit) {
    let cut = remaining.lastIndexOf('\n', limit - 1);
    if (cut >= limit / 2) {
      cut += 1;
    } else {
      cut = remaining.lastIndexOf(' ', limit);
      if (cut < limit / 2) cut = limit;
    }
    chunks.push(remaining.slice(0, cut));
    remaining = remaining.slice(cut);
  }
  if (remaining) chunks.push(remaining);
  return chunks;
}

export function approvalCard(approval: ApprovalRequest): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: '🔐 Approval requested', weight: 'Bolder', wrap: true },
      {
        type: 'FactSet',
        facts: [
          { title: 'Tool', value: approval.tool },
          { title: 'Args', value: truncate(serializeArgs(approval.args), 500) },
          { title: 'Requester', value: approval.requesterId },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        data: { action: 'approve', approvalId: approval.approvalId },
      },
      {
        type: 'Action.Execute',
        title: 'Reject',
        data: { action: 'reject', approvalId: approval.approvalId },
      },
    ],
  } as IAdaptiveCard;
}

export function toolCard(name: string, args: string): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: `🛠️ ${name}`, weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: args || '(no args)', fontType: 'Monospace', wrap: true },
    ],
  } as IAdaptiveCard;
}

export function formatToolResult(result: {
  status: string;
  error?: string;
  message?: string;
  data?: unknown;
}): string {
  if (result.status === 'error') return `Error: ${result.error ?? 'unknown error'}`;
  if (result.status === 'rejected') return result.message ?? 'Rejected.';
  if (result.status === 'pending_approval') return 'Awaiting approval.';
  return '';
}
