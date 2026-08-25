export const DEPARTMENTS = ["Vision", "Sound", "Lighting", "Staging"] as const
export type Department = string

export const RATING_LABELS: Record<number, string> = {
  0: "Not trained",
  1: "Beginner",
  2: "Developing",
  3: "Competent",
  4: "Proficient",
  5: "Expert",
}

export type EmployeeStatus = "Active" | "Inactive"

export interface Employee {
  id: string
  name: string
  department: Department
  hireDate: string
  status: EmployeeStatus
}

export interface Skill {
  id: string
  name: string
  category: string
  description: string
}

export interface Rating {
  employeeId: string
  skillId: string
  year: number
  rating: number | null
}

export interface Settings {
  targetRating: number
  logoUrl: string
}

export interface AppState {
  employees: Employee[]
  skills: Skill[]
  years: number[]
  departments: Department[]
  ratings: Rating[]
  settings: Settings
}

export type Page =
  | "howto"
  | "matrix"
  | "dashboard"
  | "employees"
  | "skills"
  | "years"
  | "departments"
  | "settings"

export type DashboardMode = "team" | "skillsets" | "individual"
