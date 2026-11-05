# V3.1 — Locate the actual Open dialog nested under the Word UIA root

V3.0 screenshot and UIA prove Ctrl+F12 opened the correct classic Open window,
with filename1148 and the actual Open split-button. It is a child Window of
the Word UIA root, not a Desktop.windows top-level window. No native option
was clicked in V3.0; all results were automation-blocked, not logical failures.

V3.1 includes descendant Window controls for dialog lookup. It also supports
an optional accelerator ampersand and keyboard fallback for filename entry.
Actual menu clicking, readonly/path/printview controls, English/Chinese and
redacted structure clipboard tests are otherwise unchanged.
