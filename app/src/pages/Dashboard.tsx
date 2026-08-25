import { useMemo, useState } from "react"
import { Scale } from "../components/Scale"
import { activeEmployees, formatAvg, rankForSkill, ratingValue, skillOverall, teamOverall } from "../logic"
import type { AppState, DashboardMode, Department } from "../types"

export function Dashboard({ state }: { state: AppState }) {
  const years = [...state.years].sort((a, b) => a - b)
  const [mode, setMode] = useState<DashboardMode>("team")
  const [fromYear, setFromYear] = useState(years[0] ?? 2023)
  const [toYear, setToYear] = useState(years[years.length - 1] ?? 2026)
  const [dept, setDept] = useState<"All" | Department>("All")
  const [skillId, setSkillId] = useState(state.skills[0]?.id ?? "")
  const [employeeId, setEmployeeId] = useState(state.employees.find((e) => e.status === "Active")?.id ?? "")

  const range = years.filter((y) => y >= fromYear && y <= toYear)
  const people = useMemo(() => activeEmployees(state, dept), [state, dept])
  const lookYear = toYear

  return (
    <section className="page">
      <h1>Dashboard</h1>
      <div className="filters">
        <label className="field">
          Look at
          <select value={mode} onChange={(e) => setMode(e.target.value as DashboardMode)}>
            <option value="team">Team</option>
            <option value="skillsets">Skillsets</option>
            <option value="individual">Individual</option>
          </select>
        </label>
        <label className="field">
          From year
          <select value={fromYear} onChange={(e) => setFromYear(Number(e.target.value))}>
            {years.map((y) => (
              <option key={y}>{y}</option>
            ))}
          </select>
        </label>
        <label className="field">
          To year
          <select value={toYear} onChange={(e) => setToYear(Number(e.target.value))}>
            {years.map((y) => (
              <option key={y}>{y}</option>
            ))}
          </select>
        </label>
        {mode === "team" && (
          <label className="field">
            Department
            <select value={dept} onChange={(e) => setDept(e.target.value as "All" | Department)}>
              <option value="All">All</option>
              {state.departments.map((d) => (
                <option key={d}>{d}</option>
              ))}
            </select>
          </label>
        )}
        {mode === "skillsets" && (
          <label className="field">
            Skill
            <select value={skillId} onChange={(e) => setSkillId(e.target.value)}>
              {state.skills.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.name}
                </option>
              ))}
            </select>
          </label>
        )}
        {mode === "individual" && (
          <label className="field">
            Employee
            <select value={employeeId} onChange={(e) => setEmployeeId(e.target.value)}>
              {state.employees.map((e) => (
                <option key={e.id} value={e.id}>
                  {e.name}
                </option>
              ))}
            </select>
          </label>
        )}
      </div>
      <Scale />

      {mode === "team" && (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Year</th>
                <th>People</th>
                <th>Team overall</th>
                {state.skills.map((s) => (
                  <th key={s.id}>{s.name}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {range.map((y) => (
                <tr key={y}>
                  <td>{y}</td>
                  <td>{people.length}</td>
                  <td>{formatAvg(teamOverall(state, y, people))}</td>
                  {state.skills.map((s) => (
                    <td key={s.id}>{formatAvg(skillOverall(state, s.id, y, people))}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {mode === "skillsets" && (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Rank</th>
                <th>Employee</th>
                <th>Dept</th>
                <th>Rating ({lookYear})</th>
              </tr>
            </thead>
            <tbody>
              {rankForSkill(state, skillId, lookYear, people).map((row, i) => (
                <tr key={row.emp.id} className={i === 0 ? "best" : undefined}>
                  <td>{i + 1}</td>
                  <td>{row.emp.name}</td>
                  <td>{row.emp.department}</td>
                  <td>{row.rating === null ? "Unrated" : row.rating}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="hint">Top row is the best candidate. Unrated people sit at the bottom. A real 0 ranks above a blank.</p>
        </div>
      )}

      {mode === "individual" && (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Skill</th>
                {range.map((y) => (
                  <th key={y}>{y}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {state.skills.map((s) => (
                <tr key={s.id}>
                  <td>{s.name}</td>
                  {range.map((y) => {
                    const v = ratingValue(state, employeeId, s.id, y)
                    return <td key={y}>{v === null ? "—" : v}</td>
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  )
}
