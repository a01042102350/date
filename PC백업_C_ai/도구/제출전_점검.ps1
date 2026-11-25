# 제출 전 점검 (읽기만 함): 시트별 수식 오류·빨간 칸·내부용 문구, 숨은 시트, 합계 대조, 남은 에러
#  2026-11-12: -YM 으로 달을 받음 (예: -YM 202610). 기본은 그 달 제출본(제출본\mesN월_제출본.xls + YYYY년M월_매출마감_결과.html)
param([int]$YM = 0, [string]$XlsPath = '', [string]$Html = '')
if (-not $YM -and $XlsPath) { $p0 = Split-Path (Split-Path $XlsPath -Parent) -Leaf; if ($p0 -eq '제출본') { $p0 = Split-Path (Split-Path (Split-Path $XlsPath -Parent) -Parent) -Leaf }; if ($p0 -match '^\d{6}$') { $YM = [int]$p0 } }
if (-not $YM) { $YM = [int](@(Get-ChildItem 'C:\ai\기초자료' -Directory | ? { $_.Name -match '^\d{6}$' -and @(Get-ChildItem $_.FullName -File -Filter 'mes*월.xls' -ErrorAction SilentlyContinue).Count } | sort Name)[-1].Name) }
$mLab = "$($YM % 100)월"; $sheetName = "mes$mLab"
if (-not $XlsPath) { $XlsPath = "C:\ai\기초자료\$YM\제출본\$($sheetName)_제출본.xls" }
if (-not $Html) { $Html = "C:\ai\기초자료\$YM\제출본\$([math]::Floor($YM / 100))년$($mLab)_매출마감_결과.html" }
"== 점검: $XlsPath"
# 엑셀을 닫을 때 오류로 꺼지면 윈도우가 '복원 모드'로 다시 켜서 파일을 읽기 전용으로 띄움 → 통합문서를 닫은 뒤 이 도구가 띄운 엑셀(PID)만 바로 끔 (2026-11-25)
if (-not ('XlPidNs.XlPid' -as [type])) { Add-Type -Namespace XlPidNs -Name XlPid -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);' }
function Close-MyExcel($app) {
  if (-not $app) { return }
  [uint32]$xp = 0; try { [void][XlPidNs.XlPid]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$xp) } catch {}
  try { foreach ($w0 in @($app.Workbooks)) { if ($w0) { $w0.Close($false) } } } catch {}
  try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if ($xp) { $pr = Get-Process -Id $xp -ErrorAction SilentlyContinue; if ($pr -and $pr.ProcessName -eq 'EXCEL' -and $pr.MainWindowHandle -eq 0) { Stop-Process -Id $xp -Force -ErrorAction SilentlyContinue } }
}
$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($XlsPath, 0, $true)
  "== 시트"
  foreach ($s in @($wb.Worksheets)) {
    $u = $s.UsedRange; $nr = $u.Row + $u.Rows.Count - 1; $nc = $u.Column + $u.Columns.Count - 1
    $vis = switch ($s.Visible) { -1 { '보임' } 0 { '숨김' } 2 { '완전 숨김' } }
    $err = 0; $red = 0; $words = @{}
    $v = $s.Range($s.Cells(1,1), $s.Cells($nr,$nc)).Value2
    for ($i = 1; $i -le $nr; $i++) { for ($j = 1; $j -le $nc; $j++) { $x = $v[$i,$j]
      if ($x -is [int]) { $err++ }   # COM은 #REF! 등 오류값을 정수로 돌려줌
      elseif ($x -is [string]) { foreach ($w in 'GPT', '클로드', 'Claude', '사용자', '#REF', '시트 참고', '테스트') { if ($x -match [regex]::Escape($w)) { $words[$w]++ } } } } }
    if ($s.Visible -eq -1) { for ($i = 2; $i -le $nr; $i++) { for ($j = 1; $j -le $nc; $j++) { if ([int]$s.Cells($i,$j).Interior.Color -eq 255) { $red++ } } } }
    "  [$($s.Name)] $vis / ${nr}행 x ${nc}열 / 수식오류 $err / 빨간 칸 $red" + $(if ($words.Count) { " / 내부용 문구: " + (($words.Keys | % { "$_ $($words[$_])" }) -join ', ') } else { '' })
  }
  "== 이름 정의(#REF 등)"; foreach ($n in @($wb.Names)) { if ("$($n.RefersTo)" -match '#REF') { "  $($n.Name) = $($n.RefersTo)" } }
  "== 합계 대조"
  $gs = $wb.Worksheets.Item("해외매출"); $gn = $gs.UsedRange.Rows.Count; "  해외매출 합계줄: 원화(회계) $($gs.Cells($gn,9).Text) / MES 금액 $($gs.Cells($gn,11).Text)"
  $ks = $wb.Worksheets.Item("국내매출"); $kn = $ks.UsedRange.Rows.Count; "  국내매출 합계줄: MES 금액 $($ks.Cells($kn,3).Text) / $mLab 매출 $($ks.Cells($kn,12).Text) / 공급가액 $($ks.Cells($kn,16).Text) / 세금계산서 차이 $($ks.Cells($kn,19).Text)"
  $ws = $wb.Worksheets.Item($sheetName); $mn = $ws.UsedRange.Rows.Count; $cA = 1; while ("$($ws.Cells(1,$cA).Text)" -ne '금액') { $cA++ }
  "  $sheetName TOTAL 금액 $($ws.Cells($mn,$cA).Text) (수식 $($ws.Cells($mn,$cA).Formula)) / 전자세금계산서 금액 $($ws.Cells($mn,$cA+1).Text)"
  $cs = $wb.Worksheets.Item("이월매출"); $cn = $cs.UsedRange.Rows.Count; "  이월매출 합계줄: 원화 $($cs.Cells($cn,8).Text)"
  "== 오류목록 분류별"
  $es = $wb.Worksheets.Item("오류목록"); $en = $es.UsedRange.Rows.Count; $byC = @{}
  for ($i = 2; $i -le $en; $i++) { $k = "$($es.Cells($i,1).Text)"; if ($k) { $byC[$k]++ } }
  $byC.Keys | % { "  $_ $($byC[$_])줄" }
  "  참고 중 '작성일 … 이번 달 비교에서 제외'(다른 달 세금계산서): " + @(2..$en | ? { "$($es.Cells($_,10).Text)" -match '이번 달 비교에서 제외' }).Count + "줄"
  $wb.Close($false)
} finally { Close-MyExcel $xl; $xl = $null }   # 엑셀 오류 종료·복원 창 방지 (2026-11-25)
"== HTML"
$h = [IO.File]::ReadAllText($Html)
foreach ($w in 'GPT', '문장 복사', '클로드', 'Claude', '사용자') { $c = ([regex]::Matches($h, [regex]::Escape($w))).Count; if ($c) { "  '$w' $c곳" } }
"  파일 이름: $([IO.Path]::GetFileName($Html))"
