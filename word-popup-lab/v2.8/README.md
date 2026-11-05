# V2.8 — Restore the Word window before visible-window discovery

V2.7 product V1.1 exited0 with stageComplete, same original path, file/Word
readonly and pageview3. The Word window was still present five seconds after
the PowerShell process exited. The test minimized it for ROT attachment and
then mistakenly filtered for visible windows before restoring it. No copy
result was obtained. This is a test-harness defect, not a product failure.

V2.8 restores/maximizes/focuses that known Word window before visible-window
discovery, as already established in V1.9. The exact same V1.1 ZIP is tested,
SHA256 8cb13e31719b0191d45a61714f5b6dd10975e1359dac078ad3a22cfbf2fa42d3.
No production file is changed. Same6keyboard/mouse/CJK/structure cases and
native A2/disabled privacy verification.
