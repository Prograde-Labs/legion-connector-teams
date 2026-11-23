export function applyGroupPrefix(
  template: string | null,
  senderName: string,
  message: string,
): string {
  if (template === null) return message;
  return template.replaceAll('{name}', senderName).replace('{message}', message);
}
