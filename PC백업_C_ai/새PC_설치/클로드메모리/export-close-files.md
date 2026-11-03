---
name: export-close-files
description: "Which files in C:\\ai\\N월마감 are the user's own vs received from 해외영업팀, and how they reconcile"
metadata:
  node_type: memory
  type: project
  originSessionId: 1f8916c0-c1b3-4627-b3ff-07277a466043
  modified: 2026-10-29T02:13:22.729Z
---

**Folder layout (from 2026-10-29, user request):** C:\ai\기초자료\YYYYMM\ holds everything for that close month: source files (N월라인, 회계, 일별마감 copy, 신코·코마츠 xlsx, 아마란스 taxData, PDFs) + mesN월.xls + ekfmsrj.html (result) + 백업\. 202609 filled, 202610 created empty. matching.md stays in C:\ai. GitHub repo mirrors: 기초자료/202609/ (results only; clone C:\ai_github\date). Older items (C:\ai\9월마감, GPT입력용, pptx, index.html…) left in place.

Older layout notes — in C:\ai\9월마감\ (monthly 수출 마감, 대신금속):
- 해외\회계 2026년 수출 (2026.09).xlsx — the user's own file (accounting). Monthly sheet e.g. `26.09(NEW)`. This is the one to check/fix.
- 해외\9월마감.xlsx, 해외\9월라인.xlsx — received from 해외영업팀 (reference, not the user's).
- 해외\2026년 수출 (2026.09).xlsx — near-copy of the 회계 file (likely 해외팀 original).
- mes\mes9월.xls — MES 출고 원장; amounts use an internal rate (~1,350 KRW/USD), so ~0.4–2.5% gaps vs BL-rate amounts are normal.

Rules the user stated (2026-10-21):
- 매출 환율 = BL일자 기준환율 (USD via `환율` sheet VLOOKUP; JPY rates typed in from 국내팀 PDF 매출내역서).
- Bergen (원화 거래, fixed 1,350) → compare KRW amount only; ignore its 소재/가공 split. For the MES 환율 column, enter Bergen's rate as-is from the 회계 file (not smbs.biz); the 1,350 may change later.
- Other FX = smbs.biz (서울외국환중개) 매매기준율 on BL date; JPY is quoted per 100 yen there.
- Compare MES vs sales per invoice by 수량 and 금액 (MES Lot = invoice no.). Folder C:\ai\기초자료\ holds 9월마감.xlsx, 9월라인.xlsx, js.pdf, 신코.pdf (국내팀 매출내역서 + 면장/CI/PL/BL).

As of 2026-10-21: 9월 invoice amounts in the 회계 file matched 9월마감 and 9월라인 exactly. See [[monthly-close-focus]].

Rules updated 2026-10-28 (mes9월 sheet, matching.md items 8–10):
- BL날짜/인보이스 source: 해외 = 9월라인 (bl 날짜, 인보이스); 국내 = 출하일 (MES 출고일), no invoice, no red; SHINKO / JS TECH(코마츠) = PDF 매출내역서. Red mark = differs from 회계 file (회계 is the only 기준; 9월마감 no longer a comparison source).
- mes9월 columns now: E 환율, F 기초자료환율, G 달러, H 기초자료달러 (PDF/9월라인 values; 달러 allocated to rows by MES 금액 so invoice sum = source).
- 면장별합계 is one sheet (면장별합계2 merged), renamed **해외매출** (item 12); 신코·코마츠 always PDF-based.
- 회계 file gray 면장 cell = always carried over to next month (exclude from 해외매출), even if BL date is in the current month (user confirmed 2026-10-28).
- Item 12 sheets: mes9월, 해외매출, **이월매출** (BL 미확정·9월라인에 없음·회계 BL 10/01 or gray → one row per invoice), **오류목록** (이월 removed; same invoice = 1 row, reasons grouped by 항목). Row-level raw list kept in very-hidden sheet 오류원본 so reruns are idempotent.
- 출고Lot typos (e.g. MES DSM2608-AL6 = AL06, an 8월 carry-over): match to the single near-identical invoice, leave the MES text as-is. User checked AL6 → no red, not an error (noted in 이월매출 참고).
- **국내매출** sheet (next to 해외매출): MES 국내 rows (excl. 신코·코마츠 PDF lots) grouped per 고객사 vs 아마란스 전자세금계산서(매출) file in 기초자료 named "*세금계산서*.xls*" (prefer "N월" prefix; header row has 거래처명·공급가액). Name match strips ㈜/(주)/주식회사/spaces, then unique prefix. Diff → 출하금액변동 (구분 국내); one-sided → red + 오류목록. User will upload 해외+국내 every month. Purpose: catch 반품·불량 재출고 that change MES 수량·금액 vs 계산서. MES has no returns; 일별마감 (2026년 일별마감list.xlsx, sheet e.g. 2609) does. 3-way per 고객사: ① MES 출고 vs 일별마감 당월 출하 (user rule 2026-10-29: 일별 < MES (−) = 반품·불량, 일별 > MES (+) = 단가 소급 → not an error; light-yellow cell, labeled "(MES와 일별마감 차이로 판단)" in 출하금액변동; only one-sided cases like MES 없음 stay 에러) ② 당월 매출 = 일별 당월 출하 + 전월 출하분 + 반품(-) + 소급 − 비고 '10월이월' (user 2026-10-29: 이월 → 이월매출 sheet like 해외, 구분=국내) ③ 세금계산서 vs 당월 매출. '불량반출' rows = 참고 only (user confirmed). Skip 소계 rows (no date & no 품번). 수정세금계산서 may be issued next month: split by **작성일** month — 작성일 in close month → included (net, noted "다음 달 발급"), 작성일 prior/next month → excluded, listed as 참고 in 오류목록 ($closeYM in script).
- MES 출하 기준 금액 ≠ 확정 금액(회계·인보이스; 국내는 앞으로 전자세금계산서) = 분류 "출하기준 금액변동", not an error → sheet **출하금액변동** (invoice 1줄, 구분 해외/국내). Sheet order: mes9월, 해외매출, 이월매출, 출하금액변동, 오류목록 (2026-10-28).
- Never show 일별마감 line codes as abbreviations: 사→사형, 자→자동, 금→금형 (user asked 2026-10-28). Unknown codes 생/구 → shown as-is with a plain explanation (meaning not yet confirmed by user).
- 메모에 Replacement = 교체품: 금액·달러·기초자료달러 0이 정상 → 0 입력, 빨간색·에러 표시 안 함 (2026-10-28).
- Results are pushed to GitHub https://github.com/a01042102350/date from clone C:\ai_github\date (files: mes9월.xls, ekfmsrj.html = HTML result, matching.md). Git at C:\Program Files\Git; push needs `GIT_TERMINAL_PROMPT=1` + `-c credential.helper=manager`.
- HTML (make_html_1028.ps1 → C:\ai\mes9월_GPT용_1028.html): 오류목록 rows have checkboxes; checked → "체크한 항목 삭제" hides them (localStorage), "되살리기" restores.
- When PDF component sum ≠ 회계 원화 (e.g. SK03 5원), 회계 (엔화×환율) is correct; show it as an error in 면장별합계.
- From now on the user names PDFs with a month prefix (e.g. 10월신코.pdf, 10월js.pdf). PDFs are scans — values must be read visually each month.
- **신코·코마츠 now come as Excel** (2026-10-29): 기초자료 "N월신코, 코마츠 지출품의서마감.xlsx"; read only sheets whose **tab is colored** (9월: "신코 222", "코마츠"); table header 인보이스/BL날짜/환율/엔화금액/사형/금형/가공/스크랩(사형)/스크랩(금형)/계, stop at '합계' row (rows below = old months). Invoice prefix "NN-" = month (> close month → 이월, < → 전월분). Missing from Excel + 회계 gray → 이월 (e.g. JS03). Excel 엔화 may have decimals (SK03 9,339,379.52) → 회계 9,339,380 is correct (user), round, and show as **알림** (yellow, new 오류목록 class between 에러 and 참고) every month.
- 15) mes9월: '전자세금계산서 금액' column right after 금액, '세금계산서 일련번호' as last column (filled from 아마란스 file: invoices grouped by 작성일, MES rows → nearest 작성일 ≥ 출고일, split by MES 금액). Sheet '전자세금계산서 금액 오류' = 국내매출 9월 매출 ≠ 공급가액 or one-sided. Sheet order: mes9월, 해외매출, 국내매출, 전자세금계산서 금액 오류, 이월매출, 출하금액변동, 오류목록.
- 아마란스 세금계산서 file = "acb2050_taxData_excel_….xlsx" (name has taxData; header 거래처명·공급가액·승인번호·대표품목). Ask user to query 작성일 1일~말일 of close month. Rules (user 2026-10-29): 일별마감 하늘색 소계(=SUM) range = 당월 매출 (= 세금계산서); rows outside a customer's 소계 = 이월 even without 비고 (only for customers that have a 소계, e.g. not STX민수). 제품 외 매출 (고철·AL재·차량매각·임차료·외주가공비, or unmatched 거래처) = not compared. 아트라스콥코 = "국·해외": 회계 면장 칸 '전자세금계산서' → in 국내매출, 당월 매출 = 회계 원화 (AT01+AT02). Name aliases: 에스앤티=SNT, 아트라스콥코코리아=ATLAS KOREA; ambiguous prefix (STX엔진㈜) → pick group whose MES 금액 = 공급가액. mes9월 row '전자세금계산서 금액' = MES 금액 when customer total matches (option 가), serials of that 작성일 kept even if several. Result 9월: all 20 customers match.
- Pending idea: monthly one-click "월마감_실행.bat" (month input) so user can run without Claude; scripts still hard-code 9월 names.
- Script: scratchpad update_1028.ps1 (session 956a71fd); PowerShell vars are case-insensitive — avoid $L/$l, $head/$HEAD clashes.
