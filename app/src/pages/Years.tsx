import { useState } from "react"
import type { AppState } from "../types"

export function Years({ state, setState }: { state: AppState; setState: (s: AppState) => void }) {
  const [year, setYear] = useState(new Date().getFullYear() + 1)

  function add() {
    if (state.years.includes(year)) return
    setState({ ...state, years: [...state.years, year].sort() })
  }

  return (
    <section className="page">
      <h1>Years</h1>
      <p className="lede">Each year is its own score grid. Add a year when a new season starts.</p>
      <div className="form-row">
        <input type="number" value={year} onChange={(e) => setYear(Number(e.target.value))} />
        <button onClick={add}>Add year</button>
      </div>
      <ul className="chip-list">
        {[...state.years]
          .sort((a, b) => b - a)
          .map((y) => (
            <li key={y} className="chip">
              {y}
            </li>
          ))}
      </ul>
    </section>
  )
}
