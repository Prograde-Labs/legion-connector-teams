import type { DetailLevel } from '@legion-collective/types';

export type TeamsTransport = 'http' | 'socket';
export type UnknownUserPolicy = 'pairing' | 'reject' | 'auto-provision' | 'fallback';
export type ToolDisplay = 'off' | 'status' | 'cards';

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
  };
}
