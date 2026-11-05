# V2.6 — End-to-end portable launcher

Runs the actual solution/v1.0 PowerShell file in a separate process, waits for
that process to exit, then independently verifies the still-visible Word
window, its original path, file attribute readonly, Document.ReadOnly, print
view3, and actual English/Chinese clipboard selection with keyboard/mouse.
Uses the fully redacted structure fixture. Reverts to a fresh native Word open
and requires the original popup. Privacy remains disabled. No user credentials.

The production script uses a new visible Word automation instance, a blank
window before Documents.Open(ReadOnly=True), and no global registry, file
attribute, OneDrive or privacy changes. Macro auto-execution is temporarily
disabled only in the newly created instance during document opening.
