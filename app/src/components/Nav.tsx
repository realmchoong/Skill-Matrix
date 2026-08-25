import type { Page } from "../types"

const LINKS: { id: Page; label: string }[] = [
  { id: "howto", label: "How to Use" },
  { id: "matrix", label: "Skill Matrix" },
  { id: "dashboard", label: "Dashboard" },
  { id: "employees", label: "Employees" },
  { id: "skills", label: "Skills" },
  { id: "years", label: "Years" },
  { id: "departments", label: "Departments" },
  { id: "settings", label: "Settings" },
]

export function Nav({ page, onChange, logoUrl }: { page: Page; onChange: (p: Page) => void; logoUrl: string }) {
  return (
    <header className="topbar">
      <div className="brand">
        {logoUrl ? <img src={logoUrl} alt="" className="logo" /> : <span className="logo-mark">SM</span>}
        <div>
          <div className="brand-title">Skill Matrix</div>
          <div className="brand-sub">Live event vision crew</div>
        </div>
      </div>
      <nav className="nav">
        {LINKS.map((link) => (
          <button
            key={link.id}
            className={page === link.id ? "nav-link active" : "nav-link"}
            onClick={() => onChange(link.id)}
          >
            {link.label}
          </button>
        ))}
      </nav>
    </header>
  )
}
