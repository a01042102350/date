param([string]$Out = "$env:TEMP\지출결의서인쇄")
# 작업용 엣지(포트 9333)에 열린 전자결재 미결문서 목록(pageCode=UBA2020)에서 문서 목록(번호·양식·첨부 수)을 읽어 목록.json으로
$S = $PSScriptRoot
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Out | Out-Null
$t = @(Get-Tabs | Where-Object { $_.url -like '*pageCode=UBA2020*' })[0]
if (-not $t) { throw '작업용 엣지에 전자결재 미결문서 목록 화면이 없습니다.' }
Connect-Tab $t
$json = Js @'
(() => {
  const li = document.querySelector('li[id^=readchk_]'); if (!li) return null;
  const k = Object.keys(li).find(k => k.startsWith('__reactInternalInstance') || k.startsWith('__reactFiber'));
  let f = li[k];
  for (let i = 0; i < 40 && f; i++, f = f.return) { const p = f.memoizedProps; if (p && Array.isArray(p.listSet)) {
    return JSON.stringify({ rows: p.listSet.map((r, idx) => ({ idx, DOC_ID: r.DOC_ID, FORM_ID: r.FORM_ID, DOC_NO: r.DOC_NO, DOC_TITLE: r.DOC_TITLE, FORM_NM: r.FORM_NM, FILE_CNT: r.FILE_CNT, USER_NM: r.USER_NM, ARRIVED_DT: r.ARRIVED_DT })) }); } }
  return null;
})()
'@
Disconnect-Tab
if (-not $json) { throw '목록을 읽지 못했습니다(화면에 문서가 없거나 화면 구조가 바뀜).' }
[IO.File]::WriteAllText("$Out\목록.json", $json, (New-Object Text.UTF8Encoding $false))
$o = $json | ConvertFrom-Json
$o.rows | ForEach-Object { "{0,2}. {1} | {2} | 첨부 {3} | {4}" -f ($_.idx + 1), $_.DOC_NO, $_.FORM_NM, $_.FILE_CNT, $_.DOC_TITLE }
"저장: $Out\목록.json"
