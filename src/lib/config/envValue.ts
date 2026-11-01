/**
 * An env value as typed in a hosting dashboard. Tolerates two paste mistakes
 * seen in production: the variable name pasted into the value
 * ("OUTLOOK_USER_ID=info@…" made every Outlook sync fail with "The requested
 * user 'OUTLOOK_USER_ID=info@…' is invalid") and surrounding quotes.
 */
export function cleanEnvValue(name: string, raw = process.env[name]): string {
  let value = (raw ?? "").trim();
  if (value.toUpperCase().startsWith(`${name.toUpperCase()}=`)) {
    value = value.slice(name.length + 1).trim();
  }
  const quoted = value.match(/^(["'])(.*)\1$/);
  if (quoted) value = quoted[2]!.trim();
  return value;
}
