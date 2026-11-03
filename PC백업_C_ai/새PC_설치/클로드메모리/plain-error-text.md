---
name: plain-error-text
description: "Error/reason text in the 마감 Excel/HTML must be plain Korean anyone can read — no codes, abbreviations or cell addresses"
metadata:
  node_type: memory
  type: feedback
  originSessionId: 956a71fd-b1ed-49c3-9f56-f1674f6036fe
  modified: 2026-10-28T23:44:47.436Z
---

오류목록·해외매출·국내매출 등 결과물에 쓰는 사유·기준값·출처 문구는 다른 사람이 한눈에 알아볼 수 있는 쉬운 문장으로 쓴다.
- 줄임말 금지: 사/자/금 → 사형·자동·금형, '사형/소재' → "MES 라인: 사형 / 가공: 없음(소재)".
- 셀 주소 금지: "일별마감 2609!R357" → "일별마감 9월 시트 357행".
- "분해 불가", "≠", "차 -5" 같은 축약 대신 무엇이 문제이고 무엇을 하면 되는지 문장으로 (예: "MES에 금액이 비어 있음 → MES 금액을 입력해 주세요").

**Why:** 2026-10-29 사용자가 "사 구 이런식이면 못알아먹어 다른사람이 한눈에 봐도 알아보기 쉽게 해야해"라고 함. 결과물은 팀이 같이 봄.
**How to apply:** 새 사유 문구를 만들 때마다 이 규칙 적용. 관련: [[reply-in-korean]], [[export-close-files]]
