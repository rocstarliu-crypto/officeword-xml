# V2.2 — Protected View page-layout fallback validation

Goal: test Word's documented ProtectedViewWindows.Open path without changing
file attributes, privacy, Office installation, or opening an editable copy.
This is a candidate fallback, not a repair of the normal read-only path.

Flow: positive writable/read-only baseline → Protected View mouse selection →
repeat launch, redacted structure, Chinese selection, keyboard/caret controls →
return to normal read-only opening and require the original popup to return.

Only synthetic text and V2.1's fully text-redacted structural fixture are used.
No source document or source metadata is uploaded. Helpers are imported from
the immutable V2.1 experiment; no previous version is modified.

Manual equivalent, if validated: File → Open → Browse → select file → arrow
beside Open → Open in Protected View. Keep Enable Editing unselected.

Primary references:
- https://learn.microsoft.com/en-us/office/vba/api/word.protectedviewwindows.open
- https://support.microsoft.com/en-us/office/collab-files/what-is-protected-view

The hosted Windows profile has no user's Microsoft/OneDrive account. The normal
path must reproduce the target warning in the same installation before a
fallback can be considered valid. Hourly Actions cap remains mandatory.
