"use client";

import {
  joinDateTimeLocal,
  quarterHourOptions,
  splitDateTimeLocal,
} from "@/lib/utils/quarterHour";

import styles from "@/components/admin/admin.module.css";

/**
 * Date picker + time dropdown in 15-minute steps. Reads and writes the same
 * `YYYY-MM-DDTHH:mm` string a `datetime-local` input did, so callers and the
 * API are unchanged. An existing off-grid time stays selectable and selected.
 */
export function QuarterHourDateTimeInput({
  value,
  onChange,
  disabled,
  label,
}: {
  value: string;
  onChange: (value: string) => void;
  disabled?: boolean;
  /** Accessible name prefix, e.g. "Start". */
  label: string;
}) {
  const { date, time } = splitDateTimeLocal(value);
  const options = quarterHourOptions(time);
  return (
    <span className={styles.quarterHourInput}>
      <input
        className={styles.input}
        type="date"
        aria-label={`${label} date`}
        value={date}
        disabled={disabled}
        onChange={(event) => onChange(joinDateTimeLocal(event.target.value, time))}
      />
      <select
        className={styles.select}
        aria-label={`${label} time`}
        value={time}
        disabled={disabled || !date}
        onChange={(event) => onChange(joinDateTimeLocal(date, event.target.value))}
      >
        {time ? null : <option value="">--:--</option>}
        {options.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    </span>
  );
}
