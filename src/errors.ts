export function describeError(error: unknown, extraSecrets: readonly string[] = []): string {
  const description = error instanceof Error ? `${error.name}: ${error.message}` : String(error);
  let sanitized = description.replace(/Bearer\s+[^\s,;"']+/gi, 'Bearer [redacted]');
  for (const secret of extraSecrets) {
    if (secret) sanitized = sanitized.replaceAll(secret, '[redacted]');
  }
  return sanitized;
}
