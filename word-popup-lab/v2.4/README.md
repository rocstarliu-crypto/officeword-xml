# V2.4 — Physical drag and caret controls

V2.3 opened Protected View correctly (same original path, file attribute readonly,
Document.ReadOnly=true, View.Type=3) but both Protected View and restored normal
mouse trials failed to copy. Screenshots show a caret only, so no negative UI
result is interpreted as a solution.

V2.4 adds a pause after mouse-down and eight intermediate physical drag steps.
It runs plain/structured/Chinese mouse trials independently, with and without
F7 caret browsing, plus independent keyboard trials. The reversal uses keyboard
selection, already proven positive in the same installation. An additional
ProtectedViewWindow.Edit transition records whether Word remains readonly.
All privacy values remain disabled. No user file or credential is uploaded.
