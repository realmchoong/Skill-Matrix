import { describe, expect, it } from "vitest"
import {
  average,
  employeeOverall,
  formatAvg,
  rankForSkill,
  sortKey,
  upsertRating,
} from "./logic"
import { buildSampleState } from "./sampleData"
import type { AppState, Employee } from "./types"

const people: Employee[] = [
  { id: "A", name: "Ada", department: "Vision", hireDate: "2020-01-01", status: "Active" },
  { id: "B", name: "Bea", department: "Vision", hireDate: "2020-01-01", status: "Active" },
  { id: "C", name: "Cam", department: "Sound", hireDate: "2020-01-01", status: "Active" },
]

function stateWith(ratings: AppState["ratings"]): AppState {
  return {
    employees: people,
    skills: [{ id: "S01", name: "Cameras", category: "Camera", description: "" }],
    years: [2026],
    departments: ["Vision", "Sound"],
    ratings,
    settings: { targetRating: 3, logoUrl: "" },
  }
}

describe("averages", () => {
  it("includes 0 and ignores blanks", () => {
    expect(average([5, 0, null])).toBe(2.5)
    expect(average([null, null])).toBeNull()
  })

  it("formats missing averages as an em dash", () => {
    expect(formatAvg(null)).toBe("—")
    expect(formatAvg(3.1)).toBe("3.10")
  })
})

describe("ranking", () => {
  it("treats a real 0 as higher than a blank", () => {
    expect(sortKey(0)).toBeGreaterThan(sortKey(null))
  })

  it("lists the strongest person first and unrated last", () => {
    const state = stateWith([
      { employeeId: "A", skillId: "S01", year: 2026, rating: 2 },
      { employeeId: "B", skillId: "S01", year: 2026, rating: 0 },
      { employeeId: "C", skillId: "S01", year: 2026, rating: null },
    ])
    const ranked = rankForSkill(state, "S01", 2026, people)
    expect(ranked.map((r) => r.emp.id)).toEqual(["A", "B", "C"])
  })
})

describe("ratings", () => {
  it("updates an existing cell instead of duplicating it", () => {
    const first = upsertRating([], { employeeId: "A", skillId: "S01", year: 2026, rating: 1 })
    const next = upsertRating(first, { employeeId: "A", skillId: "S01", year: 2026, rating: 0 })
    expect(next).toHaveLength(1)
    expect(next[0].rating).toBe(0)
  })

  it("computes overall from scored skills only", () => {
    const sample = buildSampleState()
    const alex = employeeOverall(sample, "E001", 2026)
    expect(alex).not.toBeNull()
    expect(alex).toBeGreaterThan(0)
  })
})
