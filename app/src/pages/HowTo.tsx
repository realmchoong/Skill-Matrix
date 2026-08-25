export function HowTo() {
  return (
    <section className="page">
      <h1>How to Use</h1>
      <p className="lede">
        Rate your live-event vision crew on eleven skills. Scores stay on this computer unless you export a backup.
      </p>
      <ol className="steps">
        <li>
          Open <b>Skill Matrix</b>. Pick a year tab, then score each person. Blank means not scored. <b>0</b> means not
          trained.
        </li>
        <li>
          Use <b>Dashboard</b> to look at Team (year-on-year), Skillsets (who is strongest on one skill), or Individual
          (one person across skills).
        </li>
        <li>
          Add people, skills, years, and departments from the catalog pages. Inactive employees stay in history but drop
          off the live matrix.
        </li>
        <li>
          In <b>Settings</b>, set the target rating used for gap highlighting, then export JSON if you want a backup.
        </li>
      </ol>
    </section>
  )
}
