import { useState } from "react"
import type { AppState } from "../types"

export function Skills({ state, setState }: { state: AppState; setState: (s: AppState) => void }) {
  const [form, setForm] = useState({ name: "", category: "", description: "" })

  function add() {
    if (!form.name.trim()) return
    const id = `S${String(state.skills.length + 1).padStart(2, "0")}`
    setState({
      ...state,
      skills: [...state.skills, { id, ...form, name: form.name.trim() }],
    })
    setForm({ name: "", category: "", description: "" })
  }

  return (
    <section className="page">
      <h1>Skills</h1>
      <div className="form-row">
        <input placeholder="Skill name" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
        <input placeholder="Category" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} />
        <input placeholder="Description" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} />
        <button onClick={add}>Add skill</button>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>Name</th>
              <th>Category</th>
              <th>Description</th>
            </tr>
          </thead>
          <tbody>
            {state.skills.map((skill) => (
              <tr key={skill.id}>
                <td>{skill.id}</td>
                <td>
                  <input
                    value={skill.name}
                    onChange={(e) =>
                      setState({
                        ...state,
                        skills: state.skills.map((s) => (s.id === skill.id ? { ...s, name: e.target.value } : s)),
                      })
                    }
                  />
                </td>
                <td>{skill.category}</td>
                <td>{skill.description}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}
