import { RATING_LABELS } from "../types"

export function RatingCell({
  value,
  target,
  onChange,
}: {
  value: number | null
  target?: number
  onChange: (next: number | null) => void
}) {
  const below = value !== null && target !== undefined && value < target
  return (
    <select
      className={`rating-select r-${value === null ? "empty" : value}${below ? " below-target" : ""}`}
      value={value === null ? "" : String(value)}
      onChange={(e) => onChange(e.target.value === "" ? null : Number(e.target.value))}
      aria-label="Skill rating"
    >
      <option value="">—</option>
      {[0, 1, 2, 3, 4, 5].map((n) => (
        <option key={n} value={n}>
          {n} {RATING_LABELS[n]}
        </option>
      ))}
    </select>
  )
}
