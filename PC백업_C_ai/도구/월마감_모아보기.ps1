# 월마감 모아보기: 기초자료 폴더의 달별 결과(C:\ai\기초자료\YYYYMM\ekfmsrj.html)를 한 페이지로 모음
#   결과 = C:\ai\기초자료\index.html (예전 C:\ai\index.html 9월 대사 페이지와 같은 모양)
#   표: 월별 요약 · 남은 에러 · 해외매출 · 국내매출 · 이월매출 · 출하금액변동 · 전자세금계산서 금액 오류 · 오류목록 (표마다 맨 앞 '월' 칸)
#   MES 출고 데이터(한 줄 = 한 품목)는 양이 많아 빼고, 달별 결과 페이지 링크로 연결
#   새 달 마감(월마감_HTML.ps1) 뒤에 다시 실행:  & "C:\ai\도구\월마감_모아보기.ps1"
#   결과 페이지는 읽기만 함(고치지 않음)
param([string]$Root = 'C:\ai\기초자료', [string]$Out = '')
$ErrorActionPreference = 'Stop'
if (-not $Out) { $Out = Join-Path $Root 'index.html' }

function Enc([string]$s) { [System.Net.WebUtility]::HtmlEncode($s) }
function Txt([string]$s) { ([System.Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', ' ')) -replace '\s+', ' ').Trim() }
function Num([string]$s) { $x = ((Txt $s) -replace '[^\d\.\-]', ''); if ($x -match '^-?\d+(\.\d+)?$') { [decimal]$x } else { [decimal]0 } }
function Won($d) { '{0:N0}' -f [decimal]$d }
function Href([string]$rel) { ($rel -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/' }
function Get-Sec([string]$t, [string]$id) { [regex]::Match($t, '<section class="sheet" id="' + $id + '"[\s\S]*?</section>').Value }
function Read-Table([string]$sec) {
  $heads = @([regex]::Matches([regex]::Match($sec, '<thead>[\s\S]*?</thead>').Value, '<th[^>]*>([\s\S]*?)</th>') | ForEach-Object { Txt $_.Groups[1].Value })
  $body = [regex]::Match($sec, '<tbody>([\s\S]*?)</tbody>').Groups[1].Value
  $rows = New-Object Collections.ArrayList
  foreach ($tr in [regex]::Matches($body, '<tr([^>]*)>([\s\S]*?)</tr>')) {
    if ($tr.Groups[2].Value -match 'colspan') { continue }   # '자료 없음' 같은 안내 줄
    $cells = @([regex]::Matches($tr.Groups[2].Value, '<td([^>]*)>([\s\S]*?)</td>') | ForEach-Object { [pscustomobject]@{ C = [regex]::Match($_.Groups[1].Value, 'class="([^"]*)"').Groups[1].Value; H = $_.Groups[2].Value } })
    if (@($cells | Select-Object -First 3 | Where-Object { (Txt $_.H) -in '합계', '소계', 'TOTAL' }).Count) { continue }   # 표 안 합계 줄(이월매출 맨 끝) — 건수·금액에 넣지 않음
    [void]$rows.Add([pscustomobject]@{ Tags = [regex]::Match($tr.Groups[1].Value, 'data-tags="([^"]*)"').Groups[1].Value; Cls = [regex]::Match($tr.Groups[1].Value, 'class="([^"]*)"').Groups[1].Value; Cells = $cells })
  }
  [pscustomobject]@{ Heads = $heads; Rows = $rows }
}
# 칸 이름의 'N월' → '당월', 'N+1월' → '다음 달' (달마다 칸 이름이 달라도 한 칸으로 모이게)
function Norm-Head([string]$hd, [int]$mon) { $nx = if ($mon -eq 12) { 1 } else { $mon + 1 }; $hd = [regex]::Replace($hd, "(?<!\d)$($mon)월", '당월'); [regex]::Replace($hd, "(?<!\d)$($nx)월", '다음 달') }
function Col-Of($heads, [string]$like) { for ($i = 0; $i -lt $heads.Count; $i++) { if ($heads[$i] -like $like) { return $i } }; return -1 }
function Cell-Txt($r, [int]$i) { if ($i -ge 0 -and $i -lt $r.Cells.Count) { Txt $r.Cells[$i].H } else { '' } }

$SECS = [ordered]@{ 's-sum' = '해외매출'; 's-dom' = '국내매출'; 's-carry' = '이월매출'; 's-ship' = '출하금액변동'; 's-taxerr' = '전자세금계산서 금액 오류'; 's-err' = '오류목록' }
$DATA = @{}; foreach ($id in $SECS.Keys) { $DATA[$id] = New-Object Collections.ArrayList }
$SUMM = New-Object Collections.ArrayList
$WARN = New-Object Collections.ArrayList
$dirs = @(Get-ChildItem -LiteralPath $Root -Directory | Where-Object { $_.Name -match '^\d{6}$' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'ekfmsrj.html')) } | Sort-Object Name)
if (-not $dirs) { throw "결과 페이지(ekfmsrj.html)가 있는 달 폴더가 없음: $Root" }

foreach ($dir in $dirs) {
  $ym = $dir.Name; $mon = [int]$ym.Substring(4, 2)
  $file = Get-Item -LiteralPath (Join-Path $dir.FullName 'ekfmsrj.html')
  $t = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
  $tabs = @{}
  foreach ($id in $SECS.Keys) {
    $tb = Read-Table (Get-Sec $t $id)
    $tb.Heads = @($tb.Heads | ForEach-Object { Norm-Head $_ $mon })
    $tabs[$id] = $tb
    [void]$DATA[$id].Add([pscustomobject]@{ M = $mon; YM = $ym; Tab = $tb })
  }
  $ms = [regex]::Match($t, '<b[^>]*>([\d,]+)행</b><span>MES 출고 데이터 \(빨간 칸 있는 행 ([\d,]+)\)')
  $stK = [regex]::Match($t, '<b[^>]*>([\d,]+)원</b><span>해외매출 원화\(회계\) 합계')

  $s = $tabs['s-sum']; $iK = Col-Of $s.Heads '원화금액(회계)*'
  $d = $tabs['s-dom']; $iD = Col-Of $d.Heads '당월 매출(일별마감)*'; $iTx = Col-Of $d.Heads '공급가액(세금계산서)*'
  $c = $tabs['s-carry']; $iC = Col-Of $c.Heads '원화금액(회계*'
  $e = $tabs['s-err']; $iCat = Col-Of $e.Heads '분류'
  foreach ($chk in @(@('해외매출 원화금액(회계)', $iK), @('국내 당월 매출(일별마감)', $iD), @('이월 원화금액', $iC), @('오류목록 분류', $iCat))) { if ($chk[1] -lt 0) { [void]$WARN.Add("$($mon)월: '$($chk[0])' 칸을 못 찾음") } }
  $sumK = [decimal]0; foreach ($r in $s.Rows) { $sumK += Num (Cell-Txt $r $iK) }
  $domK = [decimal]0; $taxK = [decimal]0; foreach ($r in $d.Rows) { $domK += Num (Cell-Txt $r $iD); $taxK += Num (Cell-Txt $r $iTx) }
  $carK = [decimal]0; foreach ($r in $c.Rows) { $carK += Num (Cell-Txt $r $iC) }
  $errN = @($e.Rows | Where-Object { (Cell-Txt $_ $iCat) -eq '에러' }).Count
  $refN = @($e.Rows | Where-Object { (Cell-Txt $_ $iCat) -eq '참고' }).Count
  if ($stK.Success -and (Num $stK.Groups[1].Value) -ne $sumK) { [void]$WARN.Add("$($mon)월: 해외매출 원화 합계가 결과 페이지 머리말($($stK.Groups[1].Value))과 다름($(Won $sumK))") }
  $subRel = "$ym/제출본/$($ym.Substring(0, 4))년$($mon)월_매출마감_결과.html"
  [void]$SUMM.Add([pscustomobject]@{
      M = $mon; YM = $ym
      Mes = $(if ($ms.Success) { Num $ms.Groups[1].Value } else { $null }); Red = $(if ($ms.Success) { Num $ms.Groups[2].Value } else { $null })
      SumN = $s.Rows.Count; SumK = $sumK; DomN = $d.Rows.Count; DomK = $domK; TaxK = $taxK; CarN = $c.Rows.Count; CarK = $carK
      ShipN = $tabs['s-ship'].Rows.Count; TxN = $tabs['s-taxerr'].Rows.Count; ErrAll = $e.Rows.Count; ErrN = $errN; RefN = $refN
      Link = Href "$ym/ekfmsrj.html"; Sub = $(if (Test-Path -LiteralPath (Join-Path $Root $subRel)) { Href $subRel } else { '' }); Made = $file.LastWriteTime
    })
}

# 칸 모으기: 최근 달 칸 순서를 먼저, 예전 달에만 있는 칸은 뒤에
$UNI = @{}
foreach ($id in $SECS.Keys) {
  $list = New-Object Collections.ArrayList
  foreach ($x in @($DATA[$id] | Sort-Object YM -Descending)) { foreach ($hd in $x.Tab.Heads) { if ($hd -and $hd -ne '해결' -and -not $list.Contains($hd)) { [void]$list.Add($hd) } } }
  $UNI[$id] = $list
}
function Build-Rows([string]$id, [bool]$onlyErr) {
  $sb = New-Object Text.StringBuilder; $n = 0
  foreach ($x in $DATA[$id]) {
    $idx = @{}; for ($i = 0; $i -lt $x.Tab.Heads.Count; $i++) { if (-not $idx.ContainsKey($x.Tab.Heads[$i])) { $idx[$x.Tab.Heads[$i]] = $i } }
    $ic = Col-Of $x.Tab.Heads '분류'
    foreach ($r in $x.Tab.Rows) {
      if ($onlyErr -and (Cell-Txt $r $ic) -ne '에러') { continue }
      $n++
      [void]$sb.Append('<tr data-m="').Append([string]$x.M).Append('"')
      if ($r.Tags) { [void]$sb.Append(' data-tags="').Append($r.Tags).Append('"') }
      if ($r.Cls) { [void]$sb.Append(' class="').Append($r.Cls).Append('"') }
      [void]$sb.Append('><td class="c mo">').Append([string]$x.M).Append('월</td>')
      foreach ($hd in $UNI[$id]) {
        if ($idx.ContainsKey($hd) -and $idx[$hd] -lt $r.Cells.Count) {
          $cell = $r.Cells[$idx[$hd]]
          [void]$sb.Append($(if ($cell.C) { '<td class="' + $cell.C + '">' } else { '<td>' })).Append($cell.H).Append('</td>')
        } else { [void]$sb.Append('<td></td>') }
      }
      [void]$sb.Append("</tr>`n")
    }
  }
  [pscustomobject]@{ Html = $sb.ToString(); Count = $n }
}
function Build-Head([string]$id) { '<thead><tr><th scope="col">월</th>' + (($UNI[$id] | ForEach-Object { '<th scope="col">' + (Enc $_) + '</th>' }) -join '') + '</tr></thead>' }
$monOpts = '<option value="">전체 달</option>' + (($SUMM | ForEach-Object { '<option value="' + $_.M + '">' + $_.M + '월</option>' }) -join '')
function Build-Section([string]$sid, [string]$title, [string]$desc, [string]$tableId, [string]$thead, $built, [string]$onlyLabel, [string]$onlyKind) {
  $chk = if ($onlyLabel) { '<label><input type="checkbox" class="only" data-kind="' + $onlyKind + '"> ' + $onlyLabel + '</label>' } else { '' }
  if ($built.Count -gt 0) {
    $tl = '<div class="tools" data-table="' + $tableId + '"><input type="search" placeholder="검색 (인보이스, 고객사, 사유 등)" aria-label="' + (Enc $title) + ' 검색"><select class="m" aria-label="달 고르기">' + $monOpts + '</select>' + $chk + '<span class="count"></span></div>'
    $tbl = '<div class="wrap"><table id="' + $tableId + '">' + $thead + '<tbody>' + $built.Html + '</tbody></table></div>'
  } else { $tl = ''; $tbl = '<p class="empty">해당하는 줄이 없습니다.</p>' }
  '<section class="sheet" id="' + $sid + '"><h2>' + (Enc $title) + '</h2><p class="sub">' + $desc + '</p>' + $tl + $tbl + '</section>'
}

# ---- 월별 요약 표
$tot = @{ Mes = 0; Red = 0; SumN = 0; SumK = 0; DomK = 0; TaxK = 0; CarN = 0; CarK = 0; ShipN = 0; TxN = 0; ErrAll = 0; ErrN = 0; RefN = 0 }
$sbS = New-Object Text.StringBuilder
foreach ($x in $SUMM) {
  foreach ($k in @($tot.Keys)) { if ($null -ne $x.$k) { $tot[$k] += $x.$k } }
  $links = '<a class="btn" href="' + $x.Link + '" target="_blank" rel="noopener">결과</a>' + $(if ($x.Sub) { '<a class="btn ghost" href="' + $x.Sub + '" target="_blank" rel="noopener">제출본</a>' } else { '' })
  $badE = if ($x.ErrN -gt 0) { 'num st-bad' } else { 'num' }; $badT = if ($x.TxN -gt 0) { 'num st-bad' } else { 'num' }
  [void]$sbS.Append('<tr data-m="' + $x.M + '"><td class="c mo">' + $x.M + '월</td>' +
    '<td class="num">' + $(if ($null -ne $x.Mes) { Won $x.Mes } else { '' }) + '</td><td class="num">' + $(if ($null -ne $x.Red) { Won $x.Red } else { '' }) + '</td>' +
    '<td class="num">' + (Won $x.SumN) + '</td><td class="num">' + (Won $x.SumK) + '</td>' +
    '<td class="num">' + (Won $x.DomN) + '</td><td class="num">' + (Won $x.DomK) + '</td><td class="num">' + $(if ($x.TaxK) { Won $x.TaxK } else { '' }) + '</td>' +
    '<td class="num">' + (Won $x.CarN) + '</td><td class="num">' + (Won $x.CarK) + '</td>' +
    '<td class="num">' + (Won $x.ShipN) + '</td><td class="' + $badT + '">' + (Won $x.TxN) + '</td>' +
    '<td class="num">' + (Won $x.ErrAll) + '</td><td class="' + $badE + '">' + (Won $x.ErrN) + '</td><td class="num">' + (Won $x.RefN) + '</td>' +
    '<td class="c">' + $links + '</td><td class="c">' + $x.Made.ToString('MM-dd HH:mm') + '</td></tr>' + "`n")
}
$sumHead = '<thead><tr><th scope="col">월</th><th scope="col">MES 출고 (행)</th><th scope="col">빨간 칸 있는 행</th><th scope="col">해외매출 (건)</th><th scope="col">해외매출 원화(회계) (원)</th><th scope="col">국내 고객사 (곳)</th><th scope="col">국내 당월 매출(일별마감) (원)</th><th scope="col">세금계산서 공급가액 (원)</th><th scope="col">이월매출 (건)</th><th scope="col">이월 금액 (원)</th><th scope="col">출하금액변동 (건)</th><th scope="col">전자세금계산서 금액 오류 (건)</th><th scope="col">오류목록 (줄)</th><th scope="col">에러</th><th scope="col">참고</th><th scope="col">달별 결과</th><th scope="col">만든 시각</th></tr></thead>'
$sumFoot = '<tfoot><tr><td class="c">합계</td><td class="num">' + (Won $tot.Mes) + '</td><td class="num">' + (Won $tot.Red) + '</td><td class="num">' + (Won $tot.SumN) + '</td><td class="num">' + (Won $tot.SumK) + '</td><td></td><td class="num">' + (Won $tot.DomK) + '</td><td class="num">' + (Won $tot.TaxK) + '</td><td class="num">' + (Won $tot.CarN) + '</td><td class="num">' + (Won $tot.CarK) + '</td><td class="num">' + (Won $tot.ShipN) + '</td><td class="num">' + (Won $tot.TxN) + '</td><td class="num">' + (Won $tot.ErrAll) + '</td><td class="num' + $(if ($tot.ErrN -gt 0) { ' st-bad' } else { '' }) + '">' + (Won $tot.ErrN) + '</td><td class="num">' + (Won $tot.RefN) + '</td><td></td><td></td></tr></tfoot>'

# ---- 표 만들기
$bErr = Build-Rows 's-err' $true
$bAll = @{}; foreach ($id in $SECS.Keys) { $bAll[$id] = Build-Rows $id $false }
$yr = $dirs[-1].Name.Substring(0, 4); $m1 = $SUMM[0].M; $m2 = $SUMM[-1].M
$range = if ($SUMM.Count -eq 1) { "$($m2)월" } else { "$($m1)~$($m2)월" }
$now = Get-Date -Format 'yyyy-MM-dd HH:mm'

$madeList = ($SUMM | ForEach-Object { "$($_.M)월 $($_.Made.ToString('MM-dd HH:mm'))" }) -join ' · '
$readme = @(
  @('이 페이지', "기초자료 폴더(C:\ai\기초자료\YYYYMM)에 넣은 자료로 달마다 만든 결과 페이지(ekfmsrj.html) $($SUMM.Count)개를 한 페이지로 모은 것입니다. 표마다 맨 앞 '월' 칸이 마감한 달입니다. 칸 이름의 '당월'은 그 줄의 마감 달, '다음 달'은 그다음 달입니다."),
  @('달별 결과', "MES 출고 데이터(한 줄 = 한 품목)는 양이 많아 여기서는 뺐습니다. '월별 요약'의 '결과' 버튼이나 위쪽 달 버튼을 누르면 그 달 전체 결과가 열립니다. 해결한 항목 체크도 그 달 페이지에서 합니다."),
  @('비교 기준', "회계 파일(회계 2026년 수출 (YYYY.MM).xlsx의 YY.MM(NEW) 시트)이 기준입니다. 해외 BL날짜·인보이스 = 그 달 라인 파일, 국내 = 출하일(MES 출고일), 신코·코마츠 = 매출내역서(엑셀, 예전에는 PDF)."),
  @('빨간 칸', "회계와 다르거나 자료가 없는 칸입니다. 이유는 그 줄의 '에러내용' 칸이나 오류목록의 '사유' 칸에 있습니다."),
  @('에러 / 참고', "에러 = 고쳐야 할 것. 참고 = 내부 고정환율이나 끝자리 차이처럼 정상 범위라 고칠 필요가 없는 것. '남은 에러'에는 오류목록에서 분류가 에러인 줄만 모았습니다."),
  @('이월매출', "다음 달 매출로 넘어가는 건입니다(회계 BL일이 다음 달, 라인 파일에 없음, 회계 면장번호 칸이 회색). 정상 건입니다."),
  @('출하금액변동', "MES 금액은 내부 고정환율(약 1,350원/USD)로 계산돼서 확정 금액(회계·세금계산서)과 조금 다릅니다. 에러가 아닙니다."),
  @('국내매출', "MES 출고, 일별마감, 아마란스 전자세금계산서 세 가지를 고객사별로 비교합니다. 일별마감 매출과 세금계산서 금액이 다른 고객사는 '전자세금계산서 금액 오류'에 모읍니다."),
  @('새 달 마감 순서', "C:\ai\도구에서 차례로: 1 월마감_준비 → 2 월마감_반영 → 3 월마감_HTML → 4 제출본_만들기 → 5 제출전_점검 (모두 -YM YYYYMM). 끝나면 이 페이지를 다시 만듭니다: & ""C:\ai\도구\월마감_모아보기.ps1"""),
  @('모은 결과', "달별 결과 페이지를 만든 시각: $madeList. 이 페이지를 만든 시각: $now.")
)
$readmeRows = ($readme | ForEach-Object { '<tr><td>' + (Enc $_[0]) + '</td><td>' + (Enc $_[1]) + '</td></tr>' }) -join ''

$css = @'
:root{--bg:#ffffff;--fg:#1e2a2a;--muted:#5b6b6a;--line:#d5dddc;--head:#0f3d3e;--headfg:#ffffff;--tint:#e6f0ee;--alt:#f6f8f8;--bad:#b42318;--badbg:#fdecea;--carry:#eceff1}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0f1716;--fg:#e4ecea;--muted:#9fb0ae;--line:#2a3837;--head:#1f4a4a;--headfg:#ffffff;--tint:#16302f;--alt:#141f1e;--bad:#ff8a80;--badbg:#3a1714;--carry:#1c2626}}
:root[data-theme="dark"]{--bg:#0f1716;--fg:#e4ecea;--muted:#9fb0ae;--line:#2a3837;--head:#1f4a4a;--headfg:#ffffff;--tint:#16302f;--alt:#141f1e;--bad:#ff8a80;--badbg:#3a1714;--carry:#1c2626}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.55 "Malgun Gothic","맑은 고딕",system-ui,sans-serif}
header,main{max-width:1400px;margin:0 auto;padding:0 16px}header{padding-top:28px}main{padding-bottom:40px}
h1{font-size:24px;margin:0 0 4px}h2{font-size:18px;margin:28px 0 6px}.sub{color:var(--muted);margin:0 0 16px}
.stats{display:flex;flex-wrap:wrap;gap:12px;margin:8px 0 4px}.stat{background:var(--tint);border-radius:10px;padding:10px 14px;min-width:150px}
.stat b{display:block;font-size:20px;font-variant-numeric:tabular-nums}.stat span{color:var(--muted);font-size:12px}.stat.bad b{color:var(--bad)}
.months{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:14px 0 0}.months span{color:var(--muted);font-size:12px;margin-right:4px}
.prompt{background:var(--tint);border-radius:10px;padding:12px 14px;margin:12px 0}.prompt p{margin:0 0 8px}
button{font:inherit;cursor:pointer;border:1px solid var(--line);background:var(--bg);color:var(--fg);border-radius:8px;padding:6px 12px}
button[aria-selected="true"]{background:var(--head);color:var(--headfg);border-color:var(--head)}
a.btn{display:inline-block;font-size:12px;line-height:1.4;padding:3px 10px;margin:1px 2px;border-radius:8px;border:1px solid var(--head);background:var(--head);color:var(--headfg);text-decoration:none;white-space:nowrap}
a.btn.ghost{background:var(--bg);color:var(--fg);border-color:var(--line)}a.btn:hover{opacity:.85}
nav.tabs{display:flex;flex-wrap:wrap;gap:6px;margin:20px 0 4px;position:sticky;top:0;background:var(--bg);padding:8px 0;z-index:3}
.tools{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin:8px 0}.tools input[type=search],.tools select{font:inherit;padding:6px 10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--fg);max-width:100%}
.tools input[type=search]{min-width:240px}
.count{color:var(--muted);font-size:12px}.empty{color:var(--muted);padding:12px 0}
.wrap{overflow:auto;max-height:70vh;border:1px solid var(--line);border-radius:8px}
table{border-collapse:collapse;width:max-content;min-width:100%;font-size:13px}
th,td{border-bottom:1px solid var(--line);padding:5px 8px;white-space:nowrap;text-align:left;vertical-align:top}
thead th{position:sticky;top:0;background:var(--head);color:var(--headfg);z-index:2;text-align:center;vertical-align:middle}
td{vertical-align:middle}td.c{text-align:center;font-variant-numeric:tabular-nums}td.mo{font-weight:700}
th,td{border-right:1px solid var(--line)}tbody tr:hover{background:var(--tint)}
tbody tr:nth-child(even){background:var(--alt)}tbody tr.carry{background:var(--carry);color:var(--muted)}tbody tr[hidden]{display:none}
tfoot td{position:sticky;bottom:0;background:var(--tint);font-weight:700;z-index:1}
td.num{text-align:right;font-variant-numeric:tabular-nums}td.rn{color:var(--muted);text-align:right}
td.st-bad{color:var(--bad);background:var(--badbg);font-weight:600}
td.memo{white-space:normal;min-width:140px;max-width:260px;font-size:12px;line-height:1.4;word-break:keep-all;overflow-wrap:anywhere}
#readme table{width:100%}#readme table td:first-child{font-weight:600;white-space:nowrap}#readme table td:last-child{white-space:normal;min-width:320px;word-break:keep-all;overflow-wrap:anywhere}
section[hidden]{display:none}
@media (max-width:600px){header{padding-top:16px}h1{font-size:20px}.stat{min-width:calc(50% - 6px)}.tools input[type=search]{min-width:0;width:100%}}
@media print{nav.tabs,.tools,.prompt{display:none}section[hidden]{display:block}.wrap{max-height:none;overflow:visible}}
'@
$js = @'
(function(){
var tabs=[].slice.call(document.querySelectorAll('nav.tabs button'));
function show(id){[].forEach.call(document.querySelectorAll('section.sheet'),function(s){s.hidden=s.id!==id});tabs.forEach(function(b){b.setAttribute('aria-selected',b.dataset.t===id?'true':'false')});try{localStorage.setItem('mo-tab',id)}catch(e){}}
tabs.forEach(function(b){b.addEventListener('click',function(){show(b.dataset.t)})});
var start='s-month';try{var v=localStorage.getItem('mo-tab');if(v&&document.getElementById(v))start=v}catch(e){}
var hs=decodeURIComponent(location.hash.slice(1));if(hs&&document.getElementById(hs)&&document.getElementById(hs).classList.contains('sheet'))start=hs;
show(start);
[].forEach.call(document.querySelectorAll('.tools[data-table]'),function(tl){
 var tb=document.getElementById(tl.dataset.table);if(!tb)return;var rows=[].slice.call(tb.tBodies[0].rows);
 var q=tl.querySelector('input[type=search]'),sel=tl.querySelector('select.m'),only=tl.querySelector('input.only'),cnt=tl.querySelector('.count');
 function f(){var s=(q.value||'').trim().toLowerCase(),mv=sel?sel.value:'',n=0;
  rows.forEach(function(r){var ok=true;if(mv&&r.dataset.m!==mv)ok=false;
   if(ok&&only&&only.checked){if(only.dataset.kind==='err'){ok=(','+(r.dataset.tags||'')+',').indexOf(',에러,')>=0}else{ok=!!r.querySelector('td.st-bad')}}
   if(ok&&s&&r.textContent.toLowerCase().indexOf(s)<0)ok=false;r.hidden=!ok;if(ok)n++});
  cnt.textContent=n.toLocaleString()+'줄 보임 / 전체 '+rows.length.toLocaleString()+'줄'}
 q.addEventListener('input',f);if(sel)sel.addEventListener('change',f);if(only)only.addEventListener('change',f);f()});
var cp=document.getElementById('copy');
if(cp)cp.addEventListener('click',function(){var el=document.getElementById('gpt'),t=el.textContent;
 function done(){cp.textContent='복사됨';setTimeout(function(){cp.textContent='문장 복사'},1500)}
 function old(){var r=document.createRange();r.selectNodeContents(el);var g=getSelection();g.removeAllRanges();g.addRange(r);try{document.execCommand('copy')}catch(e){}g.removeAllRanges();done()}
 if(navigator.clipboard){navigator.clipboard.writeText(t).then(done,old)}else{old()}});
})();
'@

$totErr = $tot.ErrN
$stats = '<div class="stats">' +
'<div class="stat"><b>' + $SUMM.Count + '개월</b><span>마감 결과 (' + $yr + '년 ' + $range + ')</span></div>' +
'<div class="stat"><b>' + (Won $tot.SumN) + '건</b><span>해외매출</span></div>' +
'<div class="stat"><b>' + (Won $tot.SumK) + '원</b><span>해외매출 원화(회계) 합계</span></div>' +
'<div class="stat"><b>' + (Won $tot.DomK) + '원</b><span>국내 매출(일별마감) 합계</span></div>' +
'<div class="stat"><b>' + (Won $tot.CarN) + '건</b><span>이월매출 (다음 달로)</span></div>' +
'<div class="stat' + $(if ($totErr -gt 0) { ' bad' } else { '' }) + '"><b>' + (Won $totErr) + '건</b><span>남은 에러 (오류목록 ' + (Won $tot.ErrAll) + '줄 중)</span></div></div>'
$monBtns = '<div class="months"><span>달별 결과 열기</span>' + (($SUMM | ForEach-Object { '<a class="btn ghost" href="' + $_.Link + '" target="_blank" rel="noopener">' + $_.M + '월</a>' }) -join '') + '</div>'
$gpt = "이 페이지의 'README'를 먼저 읽고, '남은 에러'와 '오류목록'의 분류 열을 기준으로 달별로 정리해줘. 비교 기준은 회계 파일이고, 신코·코마츠 행의 외화금액은 엔화야. '이월매출'(다음 달로 넘어가는 건), '출하금액변동', '참고'는 정상이니 '에러'만 정리해줘."

$nav = '<nav class="tabs" aria-label="시트"><button data-t="readme">README</button><button data-t="s-month">월별 요약 (' + $SUMM.Count + ')</button><button data-t="s-left">남은 에러 (' + $bErr.Count + ')</button>' +
  (($SECS.Keys | ForEach-Object { '<button data-t="a' + $_ + '">' + (Enc $SECS[$_]) + ' (' + (Won $bAll[$_].Count) + ')</button>' }) -join '') + '</nav>'

$DESC = @{
  's-sum'    = '달마다 해외매출(면장별). 회계 면장 중 회색(다음 달 이월)은 빠져 있음. 빨간 칸 = 회계와 다름(에러내용 칸 참고).'
  's-dom'    = '달마다 국내매출, 고객사별 1줄(신코·코마츠는 해외매출에 있음). MES 출고 ↔ 일별마감 ↔ 세금계산서 비교. 다르면 빨간 칸.'
  's-carry'  = '다음 달로 넘어가는 매출. 인보이스 1건 = 1줄, 금액은 회계 기준. 정상 건이라 오류목록에서는 뺐음.'
  's-ship'   = 'MES 출하 기준 금액과 확정 금액(회계·인보이스·세금계산서)의 차이. 내부 고정환율 때문에 생기는 정상 차이로, 에러가 아님.'
  's-taxerr' = '국내 당월 매출(일별마감)과 아마란스 전자세금계산서 공급가액이 다르거나 한쪽에만 있는 고객사.'
  's-err'    = '달마다 오류목록 전체(에러 + 참고). 같은 인보이스는 1줄로 모여 있음. 해결 체크는 그 달 결과 페이지에서.'
}
$sb = New-Object Text.StringBuilder
[void]$sb.Append('<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + $yr + '년 매출 마감 모아보기</title><style>' + $css + '</style></head><body>')
[void]$sb.Append('<header><h1>' + $yr + '년 매출 마감 대사 데이터 (' + $range + ')</h1><p class="sub">대신금속 · 기초자료 폴더 달별 결과 ' + $SUMM.Count + '개 모음 · 비교 기준 = 회계 파일 · ' + $now + ' 기준</p>' + $stats + $monBtns)
[void]$sb.Append('<div class="prompt"><p><b>GPT에 이 페이지를 줄 때 함께 붙일 문장</b></p><p id="gpt">' + (Enc $gpt) + '</p><button id="copy" type="button">문장 복사</button></div></header><main>' + $nav)
[void]$sb.Append('<section class="sheet" id="readme"><h2>README (규칙)</h2><div class="wrap"><table><thead><tr><th scope="col">구분</th><th scope="col">내용</th></tr></thead><tbody>' + $readmeRows + '</tbody></table></div></section>')
[void]$sb.Append('<section class="sheet" id="s-month"><h2>월별 요약</h2><p class="sub">달마다 결과 페이지의 숫자를 한 줄로. 에러·세금계산서 금액 오류가 있으면 빨간 칸. 금액 합계는 마감한 달을 모두 더한 값.</p><div class="wrap"><table id="t-month">' + $sumHead + '<tbody>' + $sbS.ToString() + '</tbody>' + $sumFoot + '</table></div></section>')
[void]$sb.Append((Build-Section 's-left' '남은 에러' '오류목록에서 분류가 에러인 줄만 모았습니다(고쳐야 할 것). 해결하면 그 달 자료를 고치고 월마감_반영 → 월마감_HTML → 이 페이지 순서로 다시 만듭니다.' 't-left' (Build-Head 's-err') $bErr '' ''))
foreach ($id in $SECS.Keys) {
  $ol = ''; $ok = ''
  if ($id -eq 's-err') { $ol = '에러만 보기'; $ok = 'err' } elseif ($id -in 's-sum', 's-dom') { $ol = '빨간 칸 있는 줄만 보기'; $ok = 'bad' }
  [void]$sb.Append((Build-Section ('a' + $id) $SECS[$id] (Enc $DESC[$id]) ('t' + $id) (Build-Head $id) $bAll[$id] $ol $ok))
}
[void]$sb.Append('</main><script>' + $js + '</script></body></html>')
[IO.File]::WriteAllText($Out, $sb.ToString(), (New-Object Text.UTF8Encoding $false))

"만듦: $Out ($([math]::Round((Get-Item -LiteralPath $Out).Length / 1KB)) KB)"
"달: " + (($SUMM | ForEach-Object { "$($_.M)월(해외 $($_.SumN)·국내 $($_.DomN)·이월 $($_.CarN)·에러 $($_.ErrN))" }) -join ' ')
"합계: 해외 $(Won $tot.SumN)건 $(Won $tot.SumK)원 / 국내 $(Won $tot.DomK)원 / 이월 $(Won $tot.CarN)건 / 남은 에러 $(Won $tot.ErrN)건 (오류목록 $(Won $tot.ErrAll)줄)"
foreach ($id in $SECS.Keys) { "칸 [$($SECS[$id])]: " + ($UNI[$id] -join ' | ') }
if ($WARN.Count) { '확인 필요:'; $WARN } else { '확인: 달마다 해외매출 원화 합계가 결과 페이지 머리말과 같음' }
