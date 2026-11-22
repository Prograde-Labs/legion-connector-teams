import type { ConnectorContext, ToolResult } from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import { approvalCard, type ApprovalRequest } from './render.js';

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
