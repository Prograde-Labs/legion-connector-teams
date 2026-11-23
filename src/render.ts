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
  const argsText = truncate(serializeArgs(approval.args), 500);
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: '🔐 Approval requested', weight: 'Bolder', wrap: true },
      {
        type: 'FactSet',
        facts: [
          { title: 'Tool', value: approval.tool },
          { title: 'Args', value: argsText },
          { title: 'Requester', value: approval.requesterId },
        ],
      },
      {
        type: 'Input.Text',
        id: 'reason',
        label: 'Reason (optional)',
        placeholder: 'Shown to the requester when you reject',
        isMultiline: true,
        maxLength: 500,
      },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        data: {
          action: 'approve',
          approvalId: approval.approvalId,
          tool: approval.tool,
          argsText,
          requesterId: approval.requesterId,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Reject',
        data: {
          action: 'reject',
          approvalId: approval.approvalId,
          tool: approval.tool,
          argsText,
          requesterId: approval.requesterId,
        },
      },
    ],
  } as IAdaptiveCard;
}

export interface DecidedApprovalView {
  tool?: string;
  argsText?: string;
  requesterId?: string;
  outcome: 'approved' | 'rejected' | 'error';
  message?: string;
  decidedBy?: string;
}

/** Replace an approval card with its recorded outcome, removing the actions. */
export function decidedApprovalCard(view: DecidedApprovalView): IAdaptiveCard {
  const header =
    view.outcome === 'approved'
      ? '✅ Approved'
      : view.outcome === 'rejected'
        ? '❌ Rejected'
        : '⚠️ Approval failed';
  const facts: Array<{ title: string; value: string }> = [];
  if (view.tool) facts.push({ title: 'Tool', value: view.tool });
  if (view.argsText) facts.push({ title: 'Args', value: view.argsText });
  if (view.requesterId) facts.push({ title: 'Requester', value: view.requesterId });
  if (view.decidedBy) facts.push({ title: 'Decided by', value: view.decidedBy });
  const body: unknown[] = [{ type: 'TextBlock', text: header, weight: 'Bolder', wrap: true }];
  if (facts.length) body.push({ type: 'FactSet', facts });
  if (view.message) body.push({ type: 'TextBlock', text: view.message, wrap: true });
  return { type: 'AdaptiveCard', version: '1.4', body } as IAdaptiveCard;
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
