# V2.7 — Native startup sequence in portable V1.1

V2.5 proved the route after native Word /n launch, temporary blank document,
closing the initial source, then explicit Documents.Open(ReadOnly=True).
V2.6 V1.0 product used New-Object and hit RPC_E_DISCONNECTED; not delivered.
V1.1 follows the proven native start sequence and verifies new PID and source
path before operating any Word instance. It logs the exact stage on failure.
The exact ZIP is parsed and used end to end; same six independent tests,
launcher process exit, copy, original-path readonly, page-view3, native A2,
and unchanged disabled privacy controls. Failed V1.0 is retained, never patched.
