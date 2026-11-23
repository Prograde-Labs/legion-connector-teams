export function applyGroupPrefix(
  template: string | null,
  senderName: string,
  message: string,
): string {
  if (template === null) return message;
  return template.replace(
    /\{(name|message)\}/g,
    (_match, key: string) => (key === 'name' ? senderName : message),
  );
}
