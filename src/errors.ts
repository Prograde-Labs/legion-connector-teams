export function describeError(error: unknown): string {
  const description = error instanceof Error ? `${error.name}: ${error.message}` : String(error);
  return description.replace(/Bearer\s+[^\s,;"']+/gi, 'Bearer [redacted]');
}
