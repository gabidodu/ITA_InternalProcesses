"use client";

import { useEffect, useRef, useState } from "react";

type DatePickerProps = {
  /** ISO yyyy-mm-dd, or "" for no selection. */
  value: string;
  onChange: (value: string) => void;
  /**
   * Extra blackout days on top of weekends (always blocked), keyed by ISO
   * yyyy-mm-dd with a short reason shown as the cell's tooltip — e.g. legal
   * holidays, or days already covered by one of the user's own requests.
   */
  disabledDates?: Map<string, string>;
  className?: string;
  required?: boolean;
  placeholder?: string;
};

const WEEKDAY_LABELS = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"];
const MONTH_LABELS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

function toIso(date: Date): string {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function parseIso(iso: string): Date | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  if (!match) return null;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDisplay(iso: string): string {
  const date = parseIso(iso);
  if (!date) return "";
  return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
}

function isWeekend(date: Date): boolean {
  const day = date.getDay();
  return day === 0 || day === 6;
}

/**
 * A calendar-only date picker (no free-text typing) that always displays and
 * communicates in dd/mm/yyyy regardless of the browser/OS locale — unlike a
 * native `<input type="date">`, whose displayed format silently follows the
 * user's regional settings. Weekend days and any date in `disabledDates`
 * (legal holidays, or days already covered by an existing request) are
 * greyed out and unselectable, so a leave request can never start or end on
 * a non-working or already-requested day.
 */
export function DatePicker({ value, onChange, disabledDates, className = "", required, placeholder = "dd/mm/yyyy" }: DatePickerProps) {
  const [open, setOpen] = useState(false);
  const selected = parseIso(value);
  const [viewMonth, setViewMonth] = useState(() => {
    const base = selected ?? new Date();
    return new Date(base.getFullYear(), base.getMonth(), 1);
  });
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (open) {
      const base = selected ?? new Date();
      setViewMonth(new Date(base.getFullYear(), base.getMonth(), 1));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  function isBlocked(date: Date): boolean {
    return isWeekend(date) || Boolean(disabledDates?.has(toIso(date)));
  }

  function blockedReason(date: Date): string | undefined {
    const iso = toIso(date);
    return disabledDates?.get(iso) ?? (isWeekend(date) ? "Weekend" : undefined);
  }

  function handlePick(date: Date) {
    if (isBlocked(date)) return;
    onChange(toIso(date));
    setOpen(false);
  }

  const firstOfMonth = viewMonth;
  const firstWeekday = (firstOfMonth.getDay() + 6) % 7; // Monday = 0
  const daysInMonth = new Date(firstOfMonth.getFullYear(), firstOfMonth.getMonth() + 1, 0).getDate();
  const cells: { date: Date; inMonth: boolean }[] = [];
  for (let i = 0; i < firstWeekday; i++) {
    const d = new Date(firstOfMonth.getFullYear(), firstOfMonth.getMonth(), i - firstWeekday + 1);
    cells.push({ date: d, inMonth: false });
  }
  for (let d = 1; d <= daysInMonth; d++) {
    cells.push({ date: new Date(firstOfMonth.getFullYear(), firstOfMonth.getMonth(), d), inMonth: true });
  }
  while (cells.length % 7 !== 0 || cells.length < 42) {
    const last = cells[cells.length - 1].date;
    cells.push({ date: new Date(last.getFullYear(), last.getMonth(), last.getDate() + 1), inMonth: false });
  }

  const todayIso = toIso(new Date());

  return (
    <div ref={containerRef} className="relative w-full">
      <input
        type="text"
        readOnly
        required={required}
        value={formatDisplay(value)}
        placeholder={placeholder}
        onClick={() => setOpen((v) => !v)}
        className={`w-full cursor-pointer ${className}`}
      />
      {open && (
        <div className="absolute z-20 mt-1 w-72 rounded-lg border border-line bg-surface p-3 shadow-lg">
          <div className="mb-2 flex items-center justify-between">
            <button
              type="button"
              onClick={() => setViewMonth(new Date(firstOfMonth.getFullYear(), firstOfMonth.getMonth() - 1, 1))}
              className="rounded-md px-2 py-1 text-sm text-ink-2 hover:bg-surface-2 hover:text-ink"
              aria-label="Previous month"
            >
              ‹
            </button>
            <span className="text-sm font-semibold text-ink">
              {MONTH_LABELS[firstOfMonth.getMonth()]} {firstOfMonth.getFullYear()}
            </span>
            <button
              type="button"
              onClick={() => setViewMonth(new Date(firstOfMonth.getFullYear(), firstOfMonth.getMonth() + 1, 1))}
              className="rounded-md px-2 py-1 text-sm text-ink-2 hover:bg-surface-2 hover:text-ink"
              aria-label="Next month"
            >
              ›
            </button>
          </div>
          <div className="grid grid-cols-7 gap-0.5 text-center text-[11px] font-semibold text-ink-3">
            {WEEKDAY_LABELS.map((w) => (
              <div key={w} className="py-1">
                {w}
              </div>
            ))}
          </div>
          <div className="grid grid-cols-7 gap-0.5">
            {cells.map(({ date, inMonth }, i) => {
              const iso = toIso(date);
              const blocked = isBlocked(date);
              const isSelected = iso === value;
              const isToday = iso === todayIso;
              return (
                <button
                  key={i}
                  type="button"
                  disabled={blocked}
                  title={blockedReason(date)}
                  onClick={() => handlePick(date)}
                  className={`rounded-md py-1.5 text-xs tabular-nums transition-colors ${
                    !inMonth ? "text-ink-3/40" : blocked ? "cursor-not-allowed text-ink-3/60 line-through" : "text-ink hover:bg-surface-2"
                  } ${isSelected ? "bg-accent font-bold text-accent-ink" : ""} ${
                    isToday && !isSelected ? "ring-1 ring-inset ring-line" : ""
                  }`}
                >
                  {date.getDate()}
                </button>
              );
            })}
          </div>
          <p className="mt-2 text-[11px] leading-snug text-ink-3">
            Greyed-out dates aren&apos;t available (weekend, legal holiday, or already covered by one of your requests).
          </p>
        </div>
      )}
    </div>
  );
}
