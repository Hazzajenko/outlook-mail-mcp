/**
 * Pure .env content merge: replaces matching keys in place (first occurrence
 * wins, later duplicates removed), preserves every other line verbatim, and
 * appends keys that were not present. A `null` value deletes the key instead.
 * Never touches the filesystem — the CLI layer owns the read/write.
 */
export function mergeEnvContent(
  existing: string | undefined,
  updates: Record<string, string | null>,
): string {
  const pending = new Map(Object.entries(updates));
  const out: string[] = [];

  if (existing !== undefined && existing !== "") {
    const lines = existing.split("\n");
    // A trailing newline yields one empty final element; drop it so we can
    // re-add a single trailing newline uniformly at the end.
    if (lines.at(-1) === "") lines.pop();
    const replaced = new Set<string>();

    for (const line of lines) {
      const key = [...pending.keys(), ...replaced].find((k) =>
        new RegExp(`^\\s*${escapeRegExp(k)}\\s*=`).test(line),
      );
      if (key === undefined) {
        out.push(line);
      } else if (pending.has(key)) {
        const value = pending.get(key);
        if (value !== null && value !== undefined) out.push(`${key}=${value}`);
        pending.delete(key);
        replaced.add(key);
      }
      // else: duplicate of an already-handled key — dropped.
    }
  }

  for (const [key, value] of pending) {
    if (value !== null) out.push(`${key}=${value}`);
  }

  return `${out.join("\n")}\n`;
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
