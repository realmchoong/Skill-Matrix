import { useEffect, useState } from "react"
import { Nav } from "./components/Nav"
import { Dashboard } from "./pages/Dashboard"
import { Departments } from "./pages/Departments"
import { Employees } from "./pages/Employees"
import { HowTo } from "./pages/HowTo"
import { Matrix } from "./pages/Matrix"
import { Settings } from "./pages/Settings"
import { Skills } from "./pages/Skills"
import { Years } from "./pages/Years"
import { useAppState } from "./store"
import type { Page } from "./types"

const PAGES: Page[] = ["howto", "matrix", "dashboard", "employees", "skills", "years", "departments", "settings"]

function pageFromHash(): Page {
  const hash = window.location.hash.replace("#", "") as Page
  return PAGES.includes(hash) ? hash : "howto"
}

export default function App() {
  const [state, setState] = useAppState()
  const [page, setPage] = useState<Page>(pageFromHash)

  useEffect(() => {
    const onHash = () => setPage(pageFromHash())
    window.addEventListener("hashchange", onHash)
    return () => window.removeEventListener("hashchange", onHash)
  }, [])

  function go(next: Page) {
    window.location.hash = next
    setPage(next)
  }

  return (
    <div className="app">
      <Nav page={page} onChange={go} logoUrl={state.settings.logoUrl} />
      <main>
        {page === "howto" && <HowTo />}
        {page === "matrix" && <Matrix state={state} setState={setState} />}
        {page === "dashboard" && <Dashboard state={state} />}
        {page === "employees" && <Employees state={state} setState={setState} />}
        {page === "skills" && <Skills state={state} setState={setState} />}
        {page === "years" && <Years state={state} setState={setState} />}
        {page === "departments" && <Departments state={state} setState={setState} />}
        {page === "settings" && <Settings state={state} setState={setState} />}
      </main>
    </div>
  )
}
