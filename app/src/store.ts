import { useEffect, useState } from "react"
import { buildSampleState } from "./sampleData"
import type { AppState } from "./types"

const STORAGE_KEY = "skill-matrix-v1"

export function loadState(): AppState {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return buildSampleState()
    const parsed = JSON.parse(raw) as AppState
    if (!parsed.employees || !parsed.skills || !parsed.ratings) return buildSampleState()
    return parsed
  } catch {
    return buildSampleState()
  }
}

export function saveState(state: AppState) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
}

export function useAppState() {
  const [state, setState] = useState<AppState>(loadState)

  useEffect(() => {
    saveState(state)
  }, [state])

  return [state, setState] as const
}

export function exportState(state: AppState) {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" })
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = "skill-matrix.json"
  a.click()
  URL.revokeObjectURL(url)
}

export function parseImportedState(text: string): AppState {
  const parsed = JSON.parse(text) as AppState
  if (!Array.isArray(parsed.employees) || !Array.isArray(parsed.skills) || !Array.isArray(parsed.ratings)) {
    throw new Error("That file is not a skill matrix backup.")
  }
  return parsed
}
