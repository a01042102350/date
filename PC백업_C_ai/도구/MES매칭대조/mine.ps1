# 내 1~10월 결과(기초자료\YYYYMM\ekfmsrj.html)에서 거래처×월 금액 뽑기 → mine_rows.csv
$ErrorActionPreference = 'Stop'
$W = 'C:\Users\user\AppData\Local\Temp\claude\d3df15a30ffc05d5\ac480345-177e-4ecb-8db7-e819891326b4\scratchpad\mesmatch'
function Txt([string]$s) { ([System.Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', ' ')) -replace '\s+', ' ').Trim() }
function Num([string]$s) { $x = ((Txt $s) -replace '[^\d\.\-]', ''); if ($x -match '^-?\d+(\.\d+)?$') { [decimal]$x } else { [decimal]0 } }
function Get-Sec([string]$t, [string]$id) { [regex]::Match($t, '<section class="sheet" id="' + $id + '"[\s\S]*?</section>').Value }
function Read-Table([string]$sec) {
  $heads = @([regex]::Matches([regex]::Match($sec, '<thead>[\s\S]*?</thead>').Value, '<th[^>]*>([\s\S]*?)</th>') | ForEach-Object { Txt $_.Groups[1].Value })
  $body = [regex]::Match($sec, '<tbody>([\s\S]*?)</tbody>').Groups[1].Value
  $rows = New-Object Collections.ArrayList
  foreach ($tr in [regex]::Matches($body, '<tr([^>]*)>([\s\S]*?)</tr>')) {
    if ($tr.Groups[2].Value -match 'colspan') { continue }
    $cells = @([regex]::Matches($tr.Groups[2].Value, '<td([^>]*)>([\s\S]*?)</td>') | ForEach-Object { Txt $_.Groups[2].Value })
    if (@($cells | Select-Object -First 3 | Where-Object { $_ -in '합계', '소계', 'TOTAL' }).Count) { continue }
    [void]$rows.Add($cells)
  }
  [pscustomobject]@{ Heads = $heads; Rows = $rows }
}
function Col($heads, [string]$like) { for ($i = 0; $i -lt $heads.Count; $i++) { if ($heads[$i] -like $like) { return $i } }; return -1 }
$out = New-Object Collections.ArrayList
foreach ($dir in Get-ChildItem 'C:\ai\기초자료' -Directory | Where-Object Name -match '^2026\d\d$' | Sort-Object Name) {
  $mon = [int]$dir.Name.Substring(4, 2); $t = [IO.File]::ReadAllText((Join-Path $dir.FullName 'ekfmsrj.html'), [Text.Encoding]::UTF8)
  # 해외매출: 면장 1줄
  $s = Read-Table (Get-Sec $t 's-sum'); $h = $s.Heads
  $iC = Col $h '고객사'; $iK = Col $h '원화금액(회계)*'; $iMes = Col $h 'MES 금액*'; $iInv = Col $h '인보이스 번호'; $iG = Col $h '구분'
  $lines = 'sa', 'ja', 'gu', 'g1', 'g2', 'g3', 'scs', 'scg' | ForEach-Object { $_ }
  $li = @{ sa = (Col $h '사형*'); ja = (Col $h '자동*'); gu = (Col $h '금형 주조*'); g1 = (Col $h '가공1*'); g2 = (Col $h '가공2*'); g3 = (Col $h '가공3*'); scs = (Col $h '스크랩(사형)*'); scg = (Col $h '스크랩(금형)*') }
  foreach ($r in $s.Rows) {
    $o = [ordered]@{ 월 = $mon; 출처 = '해외매출'; 고객사 = $r[$iC]; 인보이스 = $r[$iInv]; 구분 = $r[$iG]; 금액 = (Num $r[$iK]); 세금계산서 = ''; MES = (Num $r[$iMes]) }
    foreach ($k in $li.Keys) { $o[$k] = $(if ($li[$k] -ge 0) { Num $r[$li[$k]] } else { 0 }) }
    [void]$out.Add([pscustomobject]$o)
  }
  # 국내매출: 고객사 1줄
  $d = Read-Table (Get-Sec $t 's-dom'); $h = $d.Heads
  $iC = Col $h '고객사'; $iK = Col $h '*월 매출(일별마감)*'; if ($iK -lt 0) { $iK = Col $h '*매출(일별마감)*' }; $iTx = Col $h '공급가액(세금계산서)*'; $iMes = Col $h 'MES 금액*'
  $li = @{ sa = (Col $h '사형*'); ja = (Col $h '자동*'); gu = (Col $h '금형 주조*'); g1 = (Col $h '가공1*'); g2 = (Col $h '가공2*'); g3 = (Col $h '가공3*') }
  foreach ($r in $d.Rows) {
    $o = [ordered]@{ 월 = $mon; 출처 = '국내매출'; 고객사 = $r[$iC]; 인보이스 = ''; 구분 = ''; 금액 = (Num $r[$iK]); 세금계산서 = $(if ($iTx -ge 0 -and $r[$iTx] -ne '') { Num $r[$iTx] } else { '' }); MES = (Num $r[$iMes]) }
    foreach ($k in 'sa', 'ja', 'gu', 'g1', 'g2', 'g3') { $o[$k] = $(if ($li[$k] -ge 0) { Num $r[$li[$k]] } else { 0 }) }; $o['scs'] = 0; $o['scg'] = 0
    [void]$out.Add([pscustomobject]$o)
  }
  # 이월매출
  $c = Read-Table (Get-Sec $t 's-carry'); $h = $c.Heads; $iC = Col $h '고객사'; $iK = Col $h '원화금액(회계*'; $iG = Col $h '구분'; $iInv = Col $h '인보이스 번호'
  foreach ($r in $c.Rows) { [void]$out.Add([pscustomobject][ordered]@{ 월 = $mon; 출처 = '이월매출'; 고객사 = $r[$iC]; 인보이스 = $r[$iInv]; 구분 = $r[$iG]; 금액 = (Num $r[$iK]); 세금계산서 = ''; MES = 0; sa = 0; ja = 0; gu = 0; g1 = 0; g2 = 0; g3 = 0; scs = 0; scg = 0 }) }
}
$out | Export-Csv (Join-Path $W 'mine_rows.csv') -NoTypeInformation -Encoding UTF8
"줄 수: $($out.Count)"
$out | Where-Object 출처 -ne '이월매출' | Group-Object 출처, 고객사 | ForEach-Object { $g = $_.Group; "{0}`t{1}개월`t{2:N0}" -f $_.Name, ($g | Select-Object -ExpandProperty 월 -Unique).Count, ($g | Measure-Object 금액 -Sum).Sum } | Sort-Object
