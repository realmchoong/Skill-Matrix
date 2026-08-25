import { render, screen } from "@testing-library/react"
import userEvent from "@testing-library/user-event"
import { beforeEach, describe, expect, it } from "vitest"
import App from "./App"

describe("App", () => {
  beforeEach(() => {
    localStorage.clear()
    window.location.hash = ""
  })

  it("opens the skill matrix from the how-to page", async () => {
    const user = userEvent.setup()
    render(<App />)
    expect(screen.getByRole("heading", { name: "How to Use" })).toBeInTheDocument()
    await user.click(screen.getByRole("button", { name: "Skill Matrix" }))
    expect(screen.getByRole("heading", { name: "Skill Matrix" })).toBeInTheDocument()
    expect(screen.getByText("Alex Chen")).toBeInTheDocument()
  })
})
