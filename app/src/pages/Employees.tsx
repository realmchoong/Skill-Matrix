import { useState } from "react"
import type { AppState, Department, Employee } from "../types"

export function Employees({ state, setState }: { state: AppState; setState: (s: AppState) => void }) {
  const [form, setForm] = useState({ name: "", department: "Vision" as Department, hireDate: "", status: "Active" as Employee["status"] })

  function add() {
    if (!form.name.trim()) return
    const id = `E${String(state.employees.length + 1).padStart(3, "0")}`
    setState({
      ...state,
      employees: [...state.employees, { id, ...form, name: form.name.trim() }],
    })
    setForm({ name: "", department: "Vision", hireDate: "", status: "Active" })
  }

  function update(id: string, patch: Partial<Employee>) {
    setState({
      ...state,
      employees: state.employees.map((e) => (e.id === id ? { ...e, ...patch } : e)),
    })
  }

  return (
    <section className="page">
      <h1>Employees</h1>
      <div className="form-row">
        <input placeholder="Name" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
        <select value={form.department} onChange={(e) => setForm({ ...form, department: e.target.value })}>
          {state.departments.map((d) => (
            <option key={d}>{d}</option>
          ))}
        </select>
        <input type="date" value={form.hireDate} onChange={(e) => setForm({ ...form, hireDate: e.target.value })} />
        <button onClick={add}>Add employee</button>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>Name</th>
              <th>Dept</th>
              <th>Hired</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {state.employees.map((emp) => (
              <tr key={emp.id}>
                <td>{emp.id}</td>
                <td>
                  <input value={emp.name} onChange={(e) => update(emp.id, { name: e.target.value })} />
                </td>
                <td>
                  <select value={emp.department} onChange={(e) => update(emp.id, { department: e.target.value as Department })}>
                    {state.departments.map((d) => (
                      <option key={d}>{d}</option>
                    ))}
                  </select>
                </td>
                <td>
                  <input type="date" value={emp.hireDate} onChange={(e) => update(emp.id, { hireDate: e.target.value })} />
                </td>
                <td>
                  <select value={emp.status} onChange={(e) => update(emp.id, { status: e.target.value as Employee["status"] })}>
                    <option>Active</option>
                    <option>Inactive</option>
                  </select>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}
