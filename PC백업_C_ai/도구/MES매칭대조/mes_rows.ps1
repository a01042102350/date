# MES 출고 줄 대조: 내 월마감 MES(기초자료 1~10월 ekfmsrj.html s-mes) ↔ 그쪽 MES 매칭(통합)
$ErrorActionPreference = 'Stop'
$W = 'C:\Users\user\AppData\Local\Temp\claude\d3df15a30ffc05d5\ac480345-177e-4ecb-8db7-e819891326b4\scratchpad\mesmatch'
. (Join-Path $W 'common.ps1')
function T([string]$s) { ([Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', ' ')) -replace '\s+', ' ').Trim() }
function DK($v) { if ("$v" -match '^\d+(\.\d+)?$') { [DateTime]::FromOADate([double]$v).ToString('yyyy-MM-dd') } elseif ("$v" -match '(\d{4})[-./](\d{1,2})[-./](\d{1,2})') { '{0}-{1:D2}-{2:D2}' -f $Matches[1], [int]$Matches[2], [int]$Matches[3] } else { "$v" } }
# 내 MES
$my = New-Object Collections.ArrayList
foreach ($dir in Get-ChildItem 'C:\ai\기초자료' -Directory | Where-Object Name -match '^2026\d\d$' | Sort-Object Name) {
  $t = [IO.File]::ReadAllText((Join-Path $dir.FullName 'ekfmsrj.html'), [Text.Encoding]::UTF8); $s = [regex]::Match($t, '<section class="sheet" id="s-mes"[\s\S]*?</section>').Value
  $hd = @([regex]::Matches([regex]::Match($s, '<thead>[\s\S]*?</thead>').Value, '<th[^>]*>([\s\S]*?)</th>') | ForEach-Object { T $_.Groups[1].Value })
  $ix = @{}; for ($i = 0; $i -lt $hd.Count; $i++) { $k = ($hd[$i] -replace ' \(.*\)$', ''); if (-not $ix.ContainsKey($k)) { $ix[$k] = $i } }
  foreach ($tr in [regex]::Matches([regex]::Match($s, '<tbody>([\s\S]*?)</tbody>').Groups[1].Value, '<tr[^>]*>([\s\S]*?)</tr>')) {
    if ($tr.Groups[1].Value -match 'colspan') { continue }
    $c = @([regex]::Matches($tr.Groups[1].Value, '<td[^>]*>([\s\S]*?)</td>') | ForEach-Object { T $_.Groups[1].Value }); if ($c.Count -lt 10) { continue }
    [void]$my.Add([pscustomobject]@{ 월 = [int]$dir.Name.Substring(4, 2); 출고일 = (DK $c[$ix['출고일']]); 구분 = $c[$ix['구분']]; 고객사 = $c[$ix['고객사']]; 품번 = $c[$ix['품번']]; 수량 = (N $c[$ix['출고수량']]); 금액 = (N $c[$ix['금액']]); Lot = $c[$ix['출고Lot']]; 메모 = $c[$ix['메모']]; 이월 = $c[$ix['이월']]; BL = $c[$ix['BL날짜']] })
  }
}
# 그쪽 MES
$L = Get-Content (Join-Path $W 's4.tsv') -Encoding UTF8; $h = $L[3] -split "`t"; $jx = @{}; for ($i = 0; $i -lt $h.Count; $i++) { if (-not $jx.ContainsKey($h[$i])) { $jx[$h[$i]] = $i } }
$th = New-Object Collections.ArrayList
foreach ($line in $L | Select-Object -Skip 4) { $c = $line -split "`t"; if ($c[$jx['출고일']] -notmatch '^\d') { continue }
  [void]$th.Add([pscustomobject]@{ 출고일 = (DK $c[$jx['출고일']]); 구분 = $c[$jx['구분']]; 고객사 = $c[$jx['고객사']]; 품번 = $c[$jx['품번']]; 수량 = (N $c[$jx['출고수량']]); 금액 = (N $c[$jx['금액']]); Lot = $c[$jx['출고Lot']]; 메모 = $c[$jx['메모']]; 매출월 = $c[$jx['매출월']]; 매출구분 = $c[$jx['매출구분']]; 영업수량 = $c[$jx['영업 수량']]; 원화 = (N $c[$jx['원화 계']]); 원장 = $c[$jx['원장 반영']]; 비고 = $c[$jx['비고']]; 파일행 = $c[$jx['파일 행']] }) }
$my | Export-Csv (Join-Path $W 'mes_my.csv') -NoTypeInformation -Encoding UTF8; $th | Export-Csv (Join-Path $W 'mes_th.csv') -NoTypeInformation -Encoding UTF8
"내 MES $($my.Count)줄 / 그쪽 MES $($th.Count)줄"
# 그쪽: 같은 MES 줄이 여러 번 나오는지 (출고일·품번·수량·금액·Lot 같은 줄)
$dup = $th | Group-Object { "$($_.출고일)|$($_.품번)|$($_.수량)|$($_.금액)|$($_.Lot)|$($_.고객사)" } | Where-Object Count -gt 1
"그쪽 같은 MES 줄이 여러 번: $(@($dup).Count)묶음, 줄 $(($dup | Measure-Object Count -Sum).Sum)"
$dupM = $my | Group-Object { "$($_.출고일)|$($_.품번)|$($_.수량)|$($_.금액)|$($_.Lot)|$($_.고객사)" } | Where-Object Count -gt 1
"내 쪽 같은 줄이 여러 번: $(@($dupM).Count)묶음"
# 달별 줄수·수량·금액 (2026년 출고만)
'--- 출고월별 (내 / 그쪽 고유 줄)'
$thU = $th | Group-Object { "$($_.출고일)|$($_.품번)|$($_.수량)|$($_.금액)|$($_.Lot)|$($_.고객사)" } | ForEach-Object { $_.Group[0] }
for ($m = 1; $m -le 10; $m++) { $p = '2026-{0:D2}' -f $m; $a = @($my | Where-Object { $_.출고일 -like "$p*" }); $b = @($thU | Where-Object { $_.출고일 -like "$p*" }); $bAll = @($th | Where-Object { $_.출고일 -like "$p*" })
  '{0,2}월  내 {1,4}줄 수량 {2,8:N0} 금액 {3,16:N0} | 그쪽 고유 {4,4}줄(전체 {5}) 수량 {6,8:N0} 금액 {7,16:N0}' -f $m, $a.Count, ($a | Measure-Object 수량 -Sum).Sum, ($a | Measure-Object 금액 -Sum).Sum, $b.Count, $bAll.Count, ($b | Measure-Object 수량 -Sum).Sum, ($b | Measure-Object 금액 -Sum).Sum }
"내 MES 중 출고일이 그 달 밖: " + @($my | Where-Object { $_.출고일 -notlike ('2026-{0:D2}*' -f $_.월) }).Count
