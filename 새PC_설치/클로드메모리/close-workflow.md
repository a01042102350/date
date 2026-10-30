---
name: close-workflow
description: "How the user wants monthly close runs done — they instruct in chat, Claude processes on this PC and pushes results to GitHub"
metadata:
  node_type: memory
  type: feedback
  originSessionId: 956a71fd-b1ed-49c3-9f56-f1674f6036fe
  modified: 2026-10-29T02:53:57.041Z
---

사용자는 매달 이 채팅에서 말로 지시하고 폴더 위치를 알려줌 → Claude가 이 PC에서 반영하고 GitHub에 알아서 올림 (2026-10-29 결정. 클라우드/Python 전환·작업 스케줄러 자동화는 하지 않기로).

- 스크립트: C:\ai\도구\월마감_반영.ps1 (엑셀 반영), 월마감_HTML.ps1 (HTML), 엑셀없이_xlsx읽기.ps1 (Excel COM이 막혔을 때 xlsx 직접 읽기), 사용법.txt. 지금은 202609에 고정 → 새 달이면 $base/$mesPath/$backup/$closeYM, 회계 시트명(YY.MM(NEW)), 일별마감 시트(YYMM), 'mes9월' 시트명, $pdf 하드코딩 값, 46296(다음 달 1일 OADate) 등을 그 달로 바꿔서 실행.
- 자료 위치: C:\ai\기초자료\YYYYMM\ (see [[export-close-files]]). GitHub: clone C:\ai_github\date → 기초자료/YYYYMM/ 에 결과물(mesN월.xls, ekfmsrj.html)만 + matching.md. Push: `$env:GIT_TERMINAL_PROMPT='1'; git -c credential.helper=manager push`.
- 자동 동기화 (2026-10-29): 작업 스케줄러 "월마감_깃허브_동기화" 10분마다 C:\ai\도구\깃허브_자동동기화.ps1 → C:\ai\기초자료 ↔ GitHub 기초자료/ 양방향(더 새 파일만, 삭제 안 함), 원본 자료 포함. 저장소는 공개(Public) — 사용자가 "공개여도 상관없어"라고 결정. 기록: C:\ai\도구\동기화기록.txt. 제출본: C:\ai\도구\제출본_만들기.ps1 → 기초자료\YYYYMM\제출본\, 점검: 제출전_점검.ps1.
- 실행 전 mesN월.xls가 엑셀에서 열려 있으면 사용자에게 닫아 달라고 할 것. 반영 전 백업\ 에 복사.

**Why:** 사용자가 "내가 여기에 말하고 명령하고 폴더 자리 알려주면 너가 깃허브하고 알아서 올려주고 반영해줘"라고 함.
**How to apply:** 새 달 요청이 오면 위 순서로 처리, 결과·에러를 쉬운 한국어로 보고 ([[reply-in-korean]], [[plain-error-text]]).
