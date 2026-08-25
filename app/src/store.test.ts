import { afterEach, describe, expect, it } from "vitest"
import { parseImportedState } from "./store"

describe("import", () => {
  afterEach(() => {
    localStorage.clear()
  })

  it("rejects files that are not a skill matrix backup", () => {
    expect(() => parseImportedState("{}")).toThrow(/backup/)
  })

  it("accepts a complete backup object", () => {
    const parsed = parseImportedState(
      JSON.stringify({
        employees: [],
        skills: [],
        years: [2026],
        departments: ["Vision"],
        ratings: [],
        settings: { targetRating: 3, logoUrl: "" },
      }),
    )
    expect(parsed.years).toEqual([2026])
  })
})
