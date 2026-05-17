const RELATIVE_RE = /^-(\d+)([dwhm])$/;
const ISO_DATE_ONLY_RE = /^\d{4}-\d{2}-\d{2}$/;
const ISO_RE = /^\d{4}-\d{2}-\d{2}(T.*)?$/;

const MS_PER_UNIT: Record<"d" | "w" | "h" | "m", number> = {
  d: 86_400_000,
  w: 7 * 86_400_000,
  h: 3_600_000,
  m: 60_000,
};

export interface ParsedDateInput {
  date: Date;
  /** true when the input was a bare YYYY-MM-DD with no time component */
  dateOnly: boolean;
}

export function parseDateInput(input: string, now: Date = new Date()): ParsedDateInput {
  if (input.length === 0) {
    throw new Error("Empty date input");
  }

  const rel = RELATIVE_RE.exec(input);
  if (rel) {
    const n = Number(rel[1]);
    const unit = rel[2] as "d" | "w" | "h" | "m";
    return { date: new Date(now.getTime() - n * MS_PER_UNIT[unit]), dateOnly: false };
  }

  if (!ISO_RE.test(input)) {
    throw new Error(`Invalid date input: ${input}`);
  }

  const d = new Date(input);
  if (Number.isNaN(d.getTime())) {
    throw new Error(`Invalid date input: ${input}`);
  }
  return { date: d, dateOnly: ISO_DATE_ONLY_RE.test(input) };
}
