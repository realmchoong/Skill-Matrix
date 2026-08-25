import { useMemo, useState } from "react"
import { RatingCell } from "../components/RatingCell"
import { Scale } from "../components/Scale"
import { employeeOverall, formatAvg, ratingValue, upsertRating } from "../logic"
import type { AppState, Department } from "../types"

export function Matrix({ state, setState }: { state: AppState; setState: (s: AppState) => void }) {
  const years = [...state.years].sort((a, b) => b - a)
  const [year, setYear] = useState(years[0] ?? new Date().getFullYear())
  const [dept, setDept] = useState<"All" | Department>("All")

  const employees = useMemo(
    () =>
      state.employees.filter(
        (e) => e.status === "Active" && (dept === "All" || e.department === dept),
      ),
    [state.employees, dept],
  )

  return (
    <section className="page">
      <div className="page-head">
        <div>
          <h1>Skill Matrix</h1>
          <p className="lede">Type scores here. They save immediately for the selected year.</p>
        </div>
        <label className="field">
          Department
          <select value={dept} onChange={(e) => setDept(e.target.value as "All" | Department)}>
            <option value="All">All</option>
            {state.departments.map((d) => (
              <option key={d}>{d}</option>
            ))}
          </select>
        </label>
      </div>
      <Scale />
      <div className="year-tabs" role="tablist">
        {years.map((y) => (
          <button key={y} className={y === year ? "tab active" : "tab"} onClick={() => setYear(y)}>
            {y}
          </button>
        ))}
      </div>
      <div className="table-wrap">
        <table className="matrix">
          <thead>
            <tr>
              <th>Employee</th>
              <th>Dept</th>
              {state.skills.map((s) => (
                <th key={s.id} title={s.description}>
                  {s.name}
                </th>
              ))}
              <th>Overall</th>
            </tr>
          </thead>
          <tbody>
            {employees.map((emp) => (
              <tr key={emp.id}>
                <td className="sticky">{emp.name}</td>
                <td>{emp.department}</td>
                {state.skills.map((skill) => (
                  <td key={skill.id}>
                    <RatingCell
                      value={ratingValue(state, emp.id, skill.id, year)}
                      target={state.settings.targetRating}
                      onChange={(rating) =>
                        setState({
                          ...state,
                          ratings: upsertRating(state.ratings, {
                            employeeId: emp.id,
                            skillId: skill.id,
                            year,
                            rating,
                          }),
                        })
                      }
                    />
                  </td>
                ))}
                <td className="overall">{formatAvg(employeeOverall(state, emp.id, year))}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}
