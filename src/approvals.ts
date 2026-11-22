import type { ConnectorContext, ToolResult } from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import { approvalCard, sendText, type ApprovalRequest } from './render.js';

export function isAwaitingApproval(result: ToolResult): boolean {
  return result.status === 'pending_approval' ||
    (result.status === 'error' && /requires_approval|\bapproval\b/i.test(result.error ?? ''));
}

export async function sendApprovalCards(
  approvals: ApprovalRequest[],
  send: (activity: MessageActivityInput) => Promise<unknown>,
): Promise<void> {
  for (const approval of approvals) {
    await send(new MessageActivityInput().addCard('adaptive', approvalCard(approval)));
  }
}

export interface ResumedItem {
  conversationId?: string;
  response?: string;
  approvalRequests?: ApprovalRequest[];
}

/**
 * Send replies produced by resuming an approved turn, plus any successor
 * approval cards the resumed agent raised.
 */
export async function sendResumed(
  resumed: ResumedItem[],
  send: (activity: MessageActivityInput) => Promise<unknown>,
  aiLabel: boolean,
): Promise<void> {
  for (const item of resumed) {
    if (item.response) await sendText(send, aiLabel, item.response);
    if (item.approvalRequests?.length) await sendApprovalCards(item.approvalRequests, send);
  }
}

export async function handleApprovalAction(args: {
  ctx: ConnectorContext;
  userId: string;
  decision: 'approve' | 'reject';
  approvalId: string;
}): Promise<ToolResult> {
  const { result } = await args.ctx.callTool(args.userId, 'approval_response', {
    decisions: [{ approvalId: args.approvalId, decision: args.decision }],
  });
  return result;
}
