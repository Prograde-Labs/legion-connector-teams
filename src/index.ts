import type { Connector, ConnectorRuntimeDeps } from '@legion-collective/core';
import { TeamsConnector } from './connector.js';
import type { TeamsConnectorOptions } from './options.js';

export type { TeamsConnectorOptions } from './options.js';

export function connector(
  options: TeamsConnectorOptions,
  deps: ConnectorRuntimeDeps,
): Connector {
  return new TeamsConnector(options, deps);
}
