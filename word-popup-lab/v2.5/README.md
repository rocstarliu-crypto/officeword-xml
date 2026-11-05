# V2.5 — Explicit Documents.Open(ReadOnly=True)

Tests a different open path: after a fresh Word launch, create a blank window,
close the initially opened synthetic file, and use Documents.Open with explicit
ReadOnly=True. No Protected View, file conversion, registry change, or content
write. Confirm original path, file attribute and Document.ReadOnly, normal page
view, and actual clipboard content. Validate repeat launch, paced physical
mouse selection, redacted structure and Chinese selections, then return to the
standard native open path in a new process and require its popup to return.

This is an opening-path candidate, not a claimed repair of Explorer's default
open path. No user account or original document is uploaded.
