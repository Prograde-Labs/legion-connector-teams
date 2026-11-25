// src/result-card.ts
import * as ACData from 'adaptivecards-templating';
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import type { ToolResult } from '@legion-collective/core';

export interface ResultCardClicker {
  id: string;
  name?: string;
}

export interface ResultCardContext {
  status: ToolResult['status'];
  data?: unknown;
  error?: string;
  message?: string;
  tool: string;
  args?: unknown;
  clicker?: ResultCardClicker;
}

const MAX_TEXT = 2000;
const MAX_ITEMS = 10;
const MAX_FACTS = 12;

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, max - 1)}…`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function header(context: ResultCardContext): string {
  if (context.status === 'success') return `✅ ${context.tool}`;
  if (context.status === 'pending_approval') return `⏳ ${context.tool} requires approval`;
  if (context.status === 'rejected') return `🚫 ${context.tool}`;
  if (context.status === 'error' && /requires_approval/.test(context.error ?? '')) {
    return `⏳ ${context.tool} requires approval`;
  }
  return `❌ ${context.tool}`;
}

function scalarFacts(value: Record<string, unknown>): Array<{ title: string; value: string }> {
  const facts: Array<{ title: string; value: string }> = [];
  for (const [key, raw] of Object.entries(value)) {
    if (facts.length >= MAX_FACTS) break;
    if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') {
      facts.push({ title: key, value: truncate(String(raw), MAX_TEXT) });
    }
  }
  return facts;
}

function arrayBody(items: unknown[]): unknown[] {
  const body: unknown[] = [];
  for (const item of items.slice(0, MAX_ITEMS)) {
    if (isRecord(item)) {
      const facts = scalarFacts(item);
      body.push(
        facts.length
          ? { type: 'FactSet', facts }
          : {
              type: 'TextBlock',
              text: truncate(JSON.stringify(item) ?? '', MAX_TEXT),
              fontType: 'Monospace',
              wrap: true,
            },
      );
    } else {
      body.push({ type: 'TextBlock', text: `• ${truncate(String(item), MAX_TEXT)}`, wrap: true });
    }
  }
  if (items.length > MAX_ITEMS) {
    body.push({
      type: 'TextBlock',
      text: `… ${items.length - MAX_ITEMS} more`,
      isSubtle: true,
      wrap: true,
    });
  }
  return body;
}

function dataBody(data: unknown): unknown[] {
  if (data === undefined || data === null) return [];
  if (typeof data === 'string') {
    return [{ type: 'TextBlock', text: truncate(data, MAX_TEXT), wrap: true }];
  }
  if (Array.isArray(data)) return arrayBody(data);
  if (isRecord(data)) {
    const body: unknown[] = [];
    const facts = scalarFacts(data);
    if (facts.length) body.push({ type: 'FactSet', facts });
    for (const [key, value] of Object.entries(data)) {
      if (!Array.isArray(value)) continue;
      body.push({ type: 'TextBlock', text: key, weight: 'Bolder', wrap: true });
      body.push(...arrayBody(value));
    }
    if (body.length === 0) {
      body.push({
        type: 'TextBlock',
        text: truncate(JSON.stringify(data) ?? '', MAX_TEXT),
        fontType: 'Monospace',
        wrap: true,
      });
    }
    return body;
  }
  return [
    { type: 'TextBlock', text: truncate(String(data), MAX_TEXT), fontType: 'Monospace', wrap: true },
  ];
}

export function genericResultCard(context: ResultCardContext): IAdaptiveCard {
  const body: unknown[] = [
    { type: 'TextBlock', text: header(context), weight: 'Bolder', wrap: true },
  ];
  if (context.status === 'error' && /requires_approval/.test(context.error ?? '')) {
    const actor = context.clicker?.name ?? context.clicker?.id ?? 'the clicking user';
    body.push({
      type: 'TextBlock',
      text: `This tool requires approval for ${actor}.`,
      wrap: true,
    });
  } else if (context.error) {
    body.push({ type: 'TextBlock', text: truncate(context.error, MAX_TEXT), wrap: true });
  }
  if (context.message) {
    body.push({ type: 'TextBlock', text: truncate(context.message, MAX_TEXT), wrap: true });
  }
  body.push(...dataBody(context.data));
  return { type: 'AdaptiveCard', version: '1.4', body } as IAdaptiveCard;
}

export function expandResultCard(
  template: unknown,
  context: ResultCardContext,
): IAdaptiveCard | undefined {
  if (!isRecord(template)) return undefined;
  try {
    const expanded = new ACData.Template(template).expand({ $root: context });
    return isRecord(expanded) && expanded.type === 'AdaptiveCard'
      ? (expanded as unknown as IAdaptiveCard)
      : undefined;
  } catch {
    return undefined;
  }
}

export function buildResultCard(
  template: unknown,
  context: ResultCardContext,
  maxBytes: number,
): IAdaptiveCard {
  const card = expandResultCard(template, context) ?? genericResultCard(context);
  if (Buffer.byteLength(JSON.stringify(card), 'utf8') <= maxBytes) return card;
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: header(context), weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: 'Result too large to display.', wrap: true },
    ],
  } as IAdaptiveCard;
}
