import { RATING_LABELS } from "../types"

export function Scale() {
  return (
    <div className="scale">
      <span>
        <b>Blank</b> not scored
      </span>
      {Object.entries(RATING_LABELS).map(([n, label]) => (
        <span key={n}>
          <b>{n}</b> {label}
        </span>
      ))}
    </div>
  )
}
