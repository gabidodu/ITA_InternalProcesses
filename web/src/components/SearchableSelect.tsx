"use client";

import { useEffect, useRef, useState } from "react";

export type SearchableSelectOption = {
  value: string;
  label: string;
};

type SearchableSelectProps = {
  options: SearchableSelectOption[];
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  /** Visual styling applied to the input itself (border/padding/text — e.g. the shared `fieldClass`). */
  className?: string;
  /** Layout classes for the outer wrapper (e.g. `flex-1` inside a flex row). Defaults to filling the parent's width. */
  wrapperClassName?: string;
  required?: boolean;
};

/**
 * A text-input-driven combobox: type to filter, click or arrow+Enter to
 * pick. Used everywhere an employee picker would otherwise be a plain
 * `<select>` — with a large org, scrolling a native dropdown to find one
 * person is painful, so this trades it for a search-as-you-type list.
 */
export function SearchableSelect({
  options,
  value,
  onChange,
  placeholder = "Search…",
  className = "",
  wrapperClassName = "",
  required,
}: SearchableSelectProps) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [highlightedIndex, setHighlightedIndex] = useState(0);
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const selectedLabel = options.find((o) => o.value === value)?.label ?? "";

  useEffect(() => {
    if (!open) setQuery(selectedLabel);
  }, [selectedLabel, open]);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
        setQuery(selectedLabel);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [selectedLabel]);

  const filtered =
    query.trim() === "" ? options : options.filter((o) => o.label.toLowerCase().includes(query.trim().toLowerCase()));

  useEffect(() => {
    setHighlightedIndex(0);
  }, [query, open]);

  function handleSelect(option: SearchableSelectOption) {
    onChange(option.value);
    setQuery(option.label);
    setOpen(false);
  }

  return (
    <div ref={containerRef} className={`relative w-full ${wrapperClassName}`}>
      <input
        ref={inputRef}
        type="text"
        required={required}
        value={open ? query : selectedLabel}
        placeholder={placeholder}
        onFocus={() => {
          setOpen(true);
          setQuery("");
        }}
        onChange={(e) => setQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Escape") {
            setOpen(false);
            setQuery(selectedLabel);
            inputRef.current?.blur();
          } else if (e.key === "ArrowDown") {
            e.preventDefault();
            setOpen(true);
            setHighlightedIndex((i) => Math.min(i + 1, filtered.length - 1));
          } else if (e.key === "ArrowUp") {
            e.preventDefault();
            setHighlightedIndex((i) => Math.max(i - 1, 0));
          } else if (e.key === "Enter") {
            e.preventDefault();
            const picked = filtered[highlightedIndex];
            if (picked) handleSelect(picked);
          }
        }}
        className={`w-full ${className}`}
        autoComplete="off"
      />
      {open && (
        <ul className="absolute z-20 mt-1 max-h-56 w-full overflow-y-auto rounded-lg border border-line bg-surface shadow-lg">
          {filtered.length === 0 ? (
            <li className="px-3 py-2 text-sm text-ink-3">No matches</li>
          ) : (
            filtered.map((option, i) => (
              <li key={option.value || "__empty__"}>
                <button
                  type="button"
                  onClick={() => handleSelect(option)}
                  className={`block w-full px-3 py-2 text-left text-sm ${
                    i === highlightedIndex ? "bg-surface-2" : ""
                  } ${option.value === value ? "font-semibold text-ink" : "text-ink-2"} hover:bg-surface-2`}
                >
                  {option.label}
                </button>
              </li>
            ))
          )}
        </ul>
      )}
    </div>
  );
}
