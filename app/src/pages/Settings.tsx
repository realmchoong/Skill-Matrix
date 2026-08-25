import { buildSampleState } from "../sampleData"
import { exportState, parseImportedState } from "../store"
import type { AppState } from "../types"

export function Settings({ state, setState }: { state: AppState; setState: (s: AppState) => void }) {
  function importFile(file: File) {
    const reader = new FileReader()
    reader.onload = () => {
      try {
        setState(parseImportedState(String(reader.result)))
      } catch (err) {
        alert(err instanceof Error ? err.message : "Could not import that file.")
      }
    }
    reader.readAsText(file)
  }

  return (
    <section className="page">
      <h1>Settings</h1>
      <label className="field">
        Target rating
        <input
          type="number"
          min={0}
          max={5}
          value={state.settings.targetRating}
          onChange={(e) =>
            setState({ ...state, settings: { ...state.settings, targetRating: Number(e.target.value) } })
          }
        />
      </label>
      <label className="field">
        Logo URL
        <input
          value={state.settings.logoUrl}
          onChange={(e) => setState({ ...state, settings: { ...state.settings, logoUrl: e.target.value } })}
          placeholder="https://…"
        />
      </label>
      <div className="form-row">
        <button onClick={() => exportState(state)}>Export JSON backup</button>
        <label className="file-btn">
          Import JSON
          <input type="file" accept="application/json" onChange={(e) => e.target.files?.[0] && importFile(e.target.files[0])} />
        </label>
        <button
          className="danger"
          onClick={() => {
            if (confirm("Replace everything with sample data?")) setState(buildSampleState())
          }}
        >
          Reset sample data
        </button>
      </div>
    </section>
  )
}
