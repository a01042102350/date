---
name: monthly-close-focus
description: "For 대신금속 export close (수출 마감) files, only work on the current month; just point out prior-month items and always report errors found"
metadata:
  node_type: memory
  type: feedback
  originSessionId: 1f8916c0-c1b3-4627-b3ff-07277a466043
  modified: 2026-10-21T02:10:57.725Z
---

When analyzing the monthly 수출 마감 Excel files (C:\ai\N월마감\...), work only on the current closing month's data/sheets. If something belongs to the previous month (e.g. 8월 이월, 8월 오기재), just mention it — the user checks those themselves. Every time errors are found (#REF!, stale month labels, mismatched totals, invoice-number typos), report them proactively.

**Why:** User (2026-10-21) said "8월 있는건 신경안써줘도 9월껏만 너는 신경써줘 … 매번 오류가있으면 말해주면 좋겠어".

**How to apply:** Scope fixes/analysis to the current month; list prior-month findings as a short "참고(전월)" note; always end checks with an error list. Issues are fixed together one at a time with the user.
