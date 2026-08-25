import type { AppState, Employee, Rating } from "./types"

export function ratingValue(state: AppState, employeeId: string, skillId: string, year: number): number | null {
  const row = state.ratings.find(
    (r) => r.employeeId === employeeId && r.skillId === skillId && r.year === year,
  )
  return row ? row.rating : null
}

export function numericRatings(values: Array<number | null>): number[] {
  return values.filter((v): v is number => v !== null && v !== undefined)
}

export function average(values: Array<number | null>): number | null {
  const nums = numericRatings(values)
  if (nums.length === 0) return null
  return nums.reduce((a, b) => a + b, 0) / nums.length
}

export function employeeOverall(state: AppState, employeeId: string, year: number): number | null {
  return average(state.skills.map((s) => ratingValue(state, employeeId, s.id, year)))
}

export function skillOverall(state: AppState, skillId: string, year: number, employees: Employee[]): number | null {
  return average(employees.map((e) => ratingValue(state, e.id, skillId, year)))
}

export function teamOverall(state: AppState, year: number, employees: Employee[]): number | null {
  return average(employees.map((e) => employeeOverall(state, e.id, year)))
}

export function sortKey(rating: number | null): number {
  return rating === null ? -1 : rating
}

export function rankForSkill(state: AppState, skillId: string, year: number, employees: Employee[]) {
  return [...employees]
    .map((emp) => ({ emp, rating: ratingValue(state, emp.id, skillId, year) }))
    .sort((a, b) => {
      const diff = sortKey(b.rating) - sortKey(a.rating)
      if (diff !== 0) return diff
      return a.emp.name.localeCompare(b.emp.name)
    })
}

export function activeEmployees(state: AppState, department?: string): Employee[] {
  return state.employees.filter((e) => {
    if (e.status !== "Active") return false
    if (department && department !== "All") return e.department === department
    return true
  })
}

export function upsertRating(ratings: Rating[], next: Rating): Rating[] {
  const idx = ratings.findIndex(
    (r) => r.employeeId === next.employeeId && r.skillId === next.skillId && r.year === next.year,
  )
  if (idx === -1) return [...ratings, next]
  const copy = [...ratings]
  copy[idx] = next
  return copy
}

export function formatAvg(value: number | null): string {
  if (value === null) return "—"
  return value.toFixed(2)
}
