import type { AppState, Employee, Skill } from "./types"

export const SAMPLE_EMPLOYEES: Employee[] = [
  { id: "E001", name: "Alex Chen", department: "Vision", hireDate: "2021-03-15", status: "Active" },
  { id: "E002", name: "Jordan Patel", department: "Vision", hireDate: "2022-07-01", status: "Active" },
  { id: "E003", name: "Sam Rivera", department: "Sound", hireDate: "2020-11-20", status: "Active" },
  { id: "E004", name: "Taylor Kim", department: "Lighting", hireDate: "2023-01-10", status: "Active" },
  { id: "E005", name: "Morgan Lee", department: "Staging", hireDate: "2019-05-08", status: "Active" },
  { id: "E006", name: "Casey Nguyen", department: "Vision", hireDate: "2024-02-12", status: "Active" },
  { id: "E007", name: "Riley Brooks", department: "Sound", hireDate: "2021-09-30", status: "Active" },
  { id: "E008", name: "Quinn Harper", department: "Lighting", hireDate: "2022-04-18", status: "Inactive" },
]

export const SAMPLE_SKILLS: Skill[] = [
  { id: "S01", name: "Broadcast Camera Operation", category: "Camera", description: "Operate broadcast cameras on live events" },
  { id: "S02", name: "PTZ Cameras Operation", category: "Camera", description: "Operate and frame PTZ cameras" },
  { id: "S03", name: "Shading / CCU", category: "Camera", description: "Camera control unit and shading" },
  { id: "S04", name: "ATEM Vision Switching", category: "Switching", description: "Switch live vision on ATEM" },
  { id: "S05", name: "Barco Vision Switching", category: "Switching", description: "Switch live vision on Barco" },
  { id: "S06", name: "Camera Switching", category: "Switching", description: "Switch between camera sources" },
  { id: "S07", name: "Content Operation", category: "Playback", description: "Run playback and graphics content" },
  { id: "S08", name: "Systems Tech", category: "Systems", description: "Support vision systems and signal flow" },
  { id: "S09", name: "Technical Setup", category: "Systems", description: "Rig, patch, and set up vision systems" },
  { id: "S10", name: "Live Streaming", category: "Streaming", description: "Encode and monitor live streams" },
  { id: "S11", name: "Troubleshooting", category: "Support", description: "Diagnose and fix live issues" },
]

const BASE_2526: Record<string, Record<string, number>> = {
  E001: { S01: 5, S02: 4, S03: 5, S04: 4, S05: 3, S06: 4, S07: 3, S08: 4, S09: 4, S10: 3, S11: 5 },
  E002: { S01: 4, S02: 5, S03: 3, S04: 3, S05: 2, S06: 3, S07: 4, S08: 3, S09: 3, S10: 4, S11: 4 },
  E003: { S01: 2, S02: 2, S03: 1, S04: 2, S05: 1, S06: 2, S07: 5, S08: 3, S09: 3, S10: 4, S11: 4 },
  E004: { S01: 3, S02: 3, S03: 2, S04: 2, S05: 2, S06: 3, S07: 3, S08: 2, S09: 4, S10: 2, S11: 3 },
  E005: { S01: 1, S02: 1, S03: 1, S04: 1, S05: 1, S06: 2, S07: 2, S08: 4, S09: 5, S10: 2, S11: 4 },
  E006: { S01: 2, S02: 3, S03: 1, S04: 1, S05: 1, S06: 2, S07: 2, S08: 2, S09: 2, S10: 3, S11: 2 },
  E007: { S01: 2, S02: 2, S03: 1, S04: 3, S05: 2, S06: 3, S07: 4, S08: 3, S09: 3, S10: 5, S11: 4 },
  E008: { S01: 3, S02: 3, S03: 3, S04: 4, S05: 4, S06: 4, S07: 2, S08: 3, S09: 3, S10: 2, S11: 3 },
}

function hireYear(employeeId: string): number {
  const emp = SAMPLE_EMPLOYEES.find((e) => e.id === employeeId)
  return emp ? Number(emp.hireDate.slice(0, 4)) : 1900
}

function ratingFor(employeeId: string, skillId: string, year: number): number | null {
  if (year < hireYear(employeeId)) return null
  const base = BASE_2526[employeeId]?.[skillId]
  if (base == null) return null
  if (year === 2026) return base
  if (year === 2025) return Math.max(1, base - 1)
  if (year === 2024) return Math.max(1, Math.min(5, (BASE_2526[employeeId]?.[skillId] ?? 3) - 2 + ((Number(skillId.slice(1)) + year) % 2)))
  if (year === 2023) return Math.max(1, Math.min(4, (BASE_2526[employeeId]?.[skillId] ?? 2) - 2))
  return null
}

export function buildSampleState(): AppState {
  const years = [2023, 2024, 2025, 2026]
  const ratings = []
  for (const emp of SAMPLE_EMPLOYEES) {
    for (const skill of SAMPLE_SKILLS) {
      for (const year of years) {
        ratings.push({
          employeeId: emp.id,
          skillId: skill.id,
          year,
          rating: ratingFor(emp.id, skill.id, year),
        })
      }
    }
  }
  return {
    employees: SAMPLE_EMPLOYEES,
    skills: SAMPLE_SKILLS,
    years,
    departments: ["Vision", "Sound", "Lighting", "Staging"],
    ratings,
    settings: { targetRating: 3, logoUrl: "" },
  }
}
