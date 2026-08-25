import { useState } from "react"
import type { AppState, Department } from "../types"

export function Departments({ state, setState }: { state: AppState; setState: (s: AppState) => void }) {
  const [name, setName] = useState("")

  function add() {
    const next = name.trim() as Department
    if (!next || state.departments.includes(next)) return
    setState({ ...state, departments: [...state.departments, next] })
    setName("")
  }

  return (
    <section className="page">
      <h1>Departments</h1>
      <div className="form-row">
        <input placeholder="Department name" value={name} onChange={(e) => setName(e.target.value)} />
        <button onClick={add}>Add department</button>
      </div>
      <ul className="chip-list">
        {state.departments.map((d) => (
          <li key={d} className="chip">
            {d}
          </li>
        ))}
      </ul>
    </section>
  )
}
