# 최근 고친 회계 파일(기초자료\YYYYMM\회계 2026년 수출 (2026.MM).xlsx)에서 차이 난 인보이스가 어떻게 적혀 있는지 찾기 (읽기만)
$ErrorActionPreference = 'Stop'
$W = 'C:\Users\user\AppData\Local\Temp\claude\d3df15a30ffc05d5\ac480345-177e-4ecb-8db7-e819891326b4\scratchpad\mesmatch'
$Q = [ordered]@{
  '202601' = @('SM01', 'AT0', '관세', 'TOOLING', '툴링')
  '202602' = @('SM01', 'AT0', 'SK01-01', 'SK01')
  '202603' = @('EM04', 'AT0')
  '202604' = @('EM04', 'SK02', 'SK03', 'AT0')
  '202605' = @('AT0', '선급')
  '202608' = @('JS0', 'AL1', 'AL01', 'SK01-01')
  '202609' = @('JS0')
}
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 3
$mine = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$out = New-Object Collections.ArrayList
try {
  foreach ($ym in $Q.Keys) {
    $src = Get-ChildItem "C:\ai\기초자료\$ym" -Filter '회계*.xlsx' | Select-Object -First 1; $tmp = Join-Path $W "acct_$ym.xlsx"; Copy-Item -LiteralPath $src.FullName $tmp -Force
    $wb = $xl.Workbooks.Open($tmp, 0, $true); $mm = $ym.Substring(4, 2)
    $sheets = @($wb.Worksheets | ForEach-Object Name); [void]$out.Add("##### $ym  파일=$($src.Name) ($($src.LastWriteTime.ToString('MM-dd HH:mm')))  시트: $($sheets -join ', ')")
    $ws = $wb.Worksheets | Where-Object { $_.Name -match "26\.$mm" -and $_.Name -match 'NEW|갑' } | Select-Object -First 1
    if (-not $ws) { $ws = $wb.Worksheets | Where-Object { $_.Name -match "26\.$mm" } | Select-Object -First 1 }
    if (-not $ws) { [void]$out.Add('  (달 시트 못 찾음)'); $wb.Close($false); continue }
    $ur = $ws.UsedRange; $v = $ur.Value2; $r0 = $ur.Row; $c0 = $ur.Column; [void]$out.Add("  시트 [$($ws.Name)] $($ur.Address(0,0))")
    # 머리글 줄 찾기 (인보이스/면장 글자가 있는 줄)
    for ($r = 1; $r -le [Math]::Min(12, $v.GetLength(0)); $r++) { $line = (1..$v.GetLength(1) | ForEach-Object { "$($v[$r, $_])" }) -join '|'; if ($line -match '인보이스|INVOICE|면장') { [void]$out.Add("  머리($($r0 + $r - 1)행): " + ((1..[Math]::Min(26, $v.GetLength(1)) | ForEach-Object { "$($v[$r, $_])" -replace '\s+', ' ' }) -join ' | ')) } }
    foreach ($pat in $Q[$ym]) {
      $hits = 0
      for ($r = 1; $r -le $v.GetLength(0); $r++) {
        $rowTxt = (1..$v.GetLength(1) | ForEach-Object { "$($v[$r, $_])" }) -join '|'
        if ($rowTxt -match [regex]::Escape($pat)) {
          $hits++; if ($hits -gt 6) { continue }
          $ci = 0; for ($c = 1; $c -le $v.GetLength(1); $c++) { if ("$($v[$r, $c])" -match [regex]::Escape($pat)) { $ci = $c; break } }
          $cell = $ws.Cells.Item($r0 + $r - 1, $c0 + $ci - 1); $clr = '{0:X6}' -f [int]$cell.Interior.Color
          $vals = (1..[Math]::Min(26, $v.GetLength(1)) | ForEach-Object { $x = $v[$r, $_]; if ($x -is [double] -and $x -gt 40000 -and $x -lt 50000) { [DateTime]::FromOADate($x).ToString('MM-dd') } elseif ($x -is [double]) { '{0:N2}' -f $x -replace '\.00$', '' } else { "$x" -replace '\s+', ' ' } }) -join ' | '
          [void]$out.Add("  [$pat] $($r0 + $r - 1)행 (찾은 칸 색 $clr): $vals")
        }
      }
      if ($hits -eq 0) { [void]$out.Add("  [$pat] 없음") } elseif ($hits -gt 6) { [void]$out.Add("  [$pat] … 모두 $hits 줄") }
    }
    $wb.Close($false)
  }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl); $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep 2; foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $mine -contains $_.Id -and $_.MainWindowHandle -eq 0 })) { Stop-Process -Id $p.Id -Force } }
[IO.File]::WriteAllLines((Join-Path $W 'acct_find.txt'), $out, (New-Object Text.UTF8Encoding $true)); $out
