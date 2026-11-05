# V2.9 — Dismiss optional first-launch sign-in offer in test profile

V2.8 V1.1 product completed with all open-state checks true and stayed open.
The screenshot showed the test machine's optional sign-in offer blocking the
clipboard test. The new harness had omitted the Escape startup handling used
by every successful earlier native Word scenario. This was not the target
connected-experience warning. No copy claim was made from V2.8.

V2.9 waits20seconds after the PowerShell launcher exits, restores/focuses Word,
and presses Escape twice to dismiss that offer without signing in, activating,
or enabling connected experiences. Existing privacy onboarding guard is used.
The exact same production V1.1 ZIP remains unchanged (8cb13e31719b0191d45a61714f5b6dd10975e1359dac078ad3a22cfbf2fa42d3).
Six actual clipboard cases and native A2 popup/privacy controls follow.
