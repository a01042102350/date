# 제출 전 점검 (읽기만 함): 시트별 수식 오류·빨간 칸·내부용 문구, 숨은 시트, 합계 대조, 남은 에러
param([string]$XlsPath = "C:\ai\기초자료\202609\mes9월.xls", [string]$Html = "C:\ai\기초자료\202609\ekfmsrj.html")
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
  $ks = $wb.Worksheets.Item("국내매출"); $kn = $ks.UsedRange.Rows.Count; "  국내매출 합계줄: MES 금액 $($ks.Cells($kn,3).Text) / 9월 매출 $($ks.Cells($kn,12).Text) / 공급가액 $($ks.Cells($kn,16).Text) / 세금계산서 차이 $($ks.Cells($kn,19).Text)"
  $ws = $wb.Worksheets.Item("mes9월"); $mn = $ws.UsedRange.Rows.Count; $cA = 1; while ("$($ws.Cells(1,$cA).Text)" -ne '금액') { $cA++ }
  "  mes9월 TOTAL 금액 $($ws.Cells($mn,$cA).Text) (수식 $($ws.Cells($mn,$cA).Formula)) / 전자세금계산서 금액 $($ws.Cells($mn,$cA+1).Text)"
  $cs = $wb.Worksheets.Item("이월매출"); $cn = $cs.UsedRange.Rows.Count; "  이월매출 합계줄: 원화 $($cs.Cells($cn,8).Text)"
  "== 오류목록 분류별"
  $es = $wb.Worksheets.Item("오류목록"); $en = $es.UsedRange.Rows.Count; $byC = @{}
  for ($i = 2; $i -le $en; $i++) { $k = "$($es.Cells($i,1).Text)"; if ($k) { $byC[$k]++ } }
  $byC.Keys | % { "  $_ $($byC[$_])줄" }
  "  참고 중 '작성일 … 이번 달 비교에서 제외'(10월 세금계산서): " + @(2..$en | ? { "$($es.Cells($_,10).Text)" -match '이번 달 비교에서 제외' }).Count + "줄"
  $wb.Close($false)
} finally { $xl.Quit() }
"== HTML"
$h = [IO.File]::ReadAllText($Html)
foreach ($w in 'GPT', '문장 복사', '클로드', 'Claude', '사용자') { $c = ([regex]::Matches($h, [regex]::Escape($w))).Count; if ($c) { "  '$w' $c곳" } }
"  파일 이름: $([IO.Path]::GetFileName($Html))"
