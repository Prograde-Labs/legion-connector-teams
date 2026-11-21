import { describe, expect, it, vi } from 'vitest';
import { handleApprovalAction, sendApprovalCards } from './approvals.js';
import type { ConnectorContext } from '@legion-collective/core';

describe('approvals', () => {
  it('sends one card per approval request', async () => {
    const send = vi.fn(async () => undefined);
    await sendApprovalCards(
      [
        { approvalId: 'ap-1', requesterId: 'assistant', tool: 'communicate', args: {} },
        { approvalId: 'ap-2', requesterId: 'researcher', tool: 'communicate', args: {} },
      ],
      send,
    );
    expect(send).toHaveBeenCalledTimes(2);
  });

  it('calls approval_response as the clicking user', async () => {
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    const ctx = { callTool } as unknown as ConnectorContext;
    const result = await handleApprovalAction({
      ctx, userId: 'alice', decision: 'approve', approvalId: 'ap-1',
    });
    expect(callTool).toHaveBeenCalledWith('alice', 'approval_response', {
      decisions: [{ approvalId: 'ap-1', decision: 'approve' }],
    });
    expect(result.status).toBe('success');
  });
});
