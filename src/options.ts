import type { DetailLevel } from '@legion-collective/types';

export type TeamsTransport = 'http' | 'socket';
export type UnknownUserPolicy = 'pairing' | 'reject' | 'auto-provision' | 'fallback';
export type ToolDisplay = 'off' | 'status' | 'cards';
export type TeamsToolName = 'cards' | 'reactions' | 'list';
const TEAM_TOOL_NAMES: readonly TeamsToolName[] = ['cards', 'reactions', 'list'];

export interface CompactionOptions {
  agentId: string;
  instruction?: string;
}

export interface TeamsConnectorOptions {
  transport?: TeamsTransport;
  port?: number;
  publicBaseUrl?: string;
  tenantId?: string;
  stateDir?: string;
  rescanIntervalMs?: number;
  socketGeos?: string[];
  unknownUserPolicy?: UnknownUserPolicy;
  fallbackUserId?: string;
  provisionerId?: string;
  pairingParticipantId?: string;
  pairingTtlSeconds?: number;
  maxDetail?: DetailLevel;
  toolDisplay?: ToolDisplay;
  aiLabel?: boolean;
  feedback?: boolean;
  typingIntervalMs?: number;
  maxMessageChars?: number;
  groupChats?: boolean;
  groupSenderPrefix?: string | null;
  learnOutboundThread?: boolean;
  tools?: TeamsToolName[] | false;
  maxMessageRecords?: number;
  maxCardBytes?: number;
  compaction?: CompactionOptions;
}

export interface ResolvedOptions {
  transport: TeamsTransport;
  port: number;
  publicBaseUrl?: string;
  tenantId?: string;
  stateDir: string;
  rescanIntervalMs: number;
  socketGeos: string[];
  unknownUserPolicy: UnknownUserPolicy;
  fallbackUserId?: string;
  provisionerId: string;
  pairingParticipantId?: string;
  pairingTtlSeconds: number;
  maxDetail: DetailLevel;
  toolDisplay: ToolDisplay;
  aiLabel: boolean;
  feedback: boolean;
  typingIntervalMs: number;
  maxMessageChars: number;
  groupChats: boolean;
  groupSenderPrefix: string | null;
  learnOutboundThread: boolean;
  tools: Set<TeamsToolName>;
  maxMessageRecords: number;
  maxCardBytes: number;
  compaction?: CompactionOptions;
}

function resolveTools(value: TeamsConnectorOptions['tools']): Set<TeamsToolName> {
  if (value === false) return new Set();
  if (value === undefined) return new Set(TEAM_TOOL_NAMES);
  const tools = new Set<TeamsToolName>();
  for (const name of value) {
    if (!TEAM_TOOL_NAMES.includes(name)) {
      throw new Error(`unknown Teams tool '${String(name)}' (expected cards, reactions, list)`);
    }
    tools.add(name);
  }
  return tools;
}

function positiveInteger(value: number | undefined, fallback: number, field: string): number {
  if (value === undefined) return fallback;
  if (!Number.isInteger(value) || value <= 0) {
    throw new Error(`${field} must be a positive integer`);
  }
  return value;
}

function resolveCompaction(
  value: TeamsConnectorOptions['compaction'],
): CompactionOptions | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error('compaction must be an object with an agentId');
  }
  if (typeof value.agentId !== 'string' || !value.agentId.trim()) {
    throw new Error('compaction.agentId must be a non-empty string');
  }
  if (value.instruction !== undefined && typeof value.instruction !== 'string') {
    throw new Error('compaction.instruction must be a string');
  }
  return {
    agentId: value.agentId.trim(),
    ...(value.instruction === undefined ? {} : { instruction: value.instruction }),
  };
}

export function resolveOptions(options: TeamsConnectorOptions): ResolvedOptions {
  const workspace = process.env.LEGION_WORKSPACE ?? process.cwd();
  const groupSenderPrefix = options.groupSenderPrefix === undefined
    ? '[{name}]: {message}'
    : options.groupSenderPrefix;
  if (groupSenderPrefix !== null && !groupSenderPrefix.includes('{message}')) {
    throw new Error('groupSenderPrefix must contain {message}');
  }
  return {
    transport: options.transport ?? 'http',
    port: options.port ?? 3978,
    publicBaseUrl: options.publicBaseUrl,
    tenantId: options.tenantId,
    stateDir: options.stateDir ?? `${workspace}/.legion/connectors/teams`,
    rescanIntervalMs: options.rescanIntervalMs ?? 60_000,
    socketGeos: options.socketGeos ?? ['amer', 'emea', 'apac'],
    unknownUserPolicy: options.unknownUserPolicy ?? 'pairing',
    fallbackUserId: options.fallbackUserId,
    provisionerId: options.provisionerId ?? 'operator',
    pairingParticipantId: options.pairingParticipantId,
    pairingTtlSeconds: options.pairingTtlSeconds ?? 600,
    maxDetail: options.maxDetail ?? 'standard',
    toolDisplay: options.toolDisplay ?? 'status',
    aiLabel: options.aiLabel ?? true,
    feedback: options.feedback ?? false,
    typingIntervalMs: options.typingIntervalMs ?? 4000,
    maxMessageChars: options.maxMessageChars ?? 20000,
    groupChats: options.groupChats ?? true,
    groupSenderPrefix,
    learnOutboundThread: options.learnOutboundThread ?? true,
    tools: resolveTools(options.tools),
    maxMessageRecords: positiveInteger(options.maxMessageRecords, 2000, 'maxMessageRecords'),
    maxCardBytes: positiveInteger(options.maxCardBytes, 28000, 'maxCardBytes'),
    compaction: resolveCompaction(options.compaction),
  };
}
