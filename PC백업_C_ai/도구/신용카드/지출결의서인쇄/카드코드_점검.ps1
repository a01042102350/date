param(
  [string]$File = ''    # 검사할 업로드 엑셀. 비우면 바탕화면\지출결의서 의 가장 최근 것
)
# 만들어진 자동전표 엑셀에서 카드 거래처코드가 규칙대로 들어갔는지 검사한다
#   미지급금(2530000) 거래처 / 부가세줄(1350000)의 결제카드(CT_NB)
#   둘 다 카드사 코드여야 한다: 기업 99706 / 경남 99709 / 농협 99708 / 현대 99903 / 하나 99905 / 국민 99701
$ErrorActionPreference = 'Stop'
$CODE = @{ '기업' = '99706'; '경남' = '99709'; '농협' = '99708'; '현대' = '99903'; '하나' = '99905'; '국민' = '99701' }

if (-not $File) {
  $File = (Get-ChildItem "$env:USERPROFILE\Desktop\지출결의서\결과" -Filter '지출결의_카드자동전표_*.xlsx' |
      Where-Object { $_.Name -notlike '*_검토*' -and $_.Name -notlike '*_분개확인*' } |
      Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}
if (-not $File -or -not (Test-Path $File)) { Write-Host '검사할 파일이 없습니다.' -ForegroundColor Red; exit 1 }
"검사: $File"
''

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($File, 0, $true)
$ws = $wb.Worksheets.Item(1); $ur = $ws.UsedRange
$pos = @{}
for ($c = 1; $c -le $ur.Columns.Count; $c++) { $k = $ur.Cells(2, $c).Text; if ($k -and -not $pos.ContainsKey($k)) { $pos[$k] = $c } }

$checked = 0; $bad = 0
for ($r = 4; $r -le $ur.Rows.Count; $r++) {
  if (-not $ur.Cells($r, 1).Text) { continue }
  $acct = $ur.Cells($r, $pos['ACCT_CD']).Text
  $trcd = $ur.Cells($r, $pos['TR_CD']).Text
  $trnm = $ur.Cells($r, $pos['TR_NM']).Text
  $ctnb = if ($pos.ContainsKey('CT_NB')) { $ur.Cells($r, $pos['CT_NB']).Text } else { '' }
  $rmk = $ur.Cells($r, $pos['RMK_DC']).Text
  $sq = $ur.Cells($r, $pos['MENU_SQ']).Text
  $cardNm = ''
  if ($rmk -match '_#([가-힣]+)') { $cardNm = $matches[1] }
  $want = if ($cardNm -and $CODE.ContainsKey($cardNm)) { $CODE[$cardNm] } else { '' }

  if ($acct -eq '2530000') {
    $checked++
    if (-not $want) { $bad++; Write-Host ("  전표{0} 미지급금: 적요에서 카드사를 못 읽음 (적요 끝이 _#기업1234 처럼 돼 있어야 함)" -f $sq) -ForegroundColor Yellow }
    elseif ($trcd -ne $want) { $bad++; Write-Host ("  전표{0} 미지급금: {1}카드인데 {2}({3}) 넣음 -> {4} 여야 함" -f $sq, $cardNm, $trcd, $trnm, $want) -ForegroundColor Red }
  }
  if ($acct -eq '1350000') {
    $checked++
    if (-not $ctnb) { $bad++; Write-Host ("  전표{0} 부가세 결제카드 칸이 비었음 -> {1} 여야 함" -f $sq, $want) -ForegroundColor Red }
    elseif ($want -and $ctnb -ne $want) { $bad++; Write-Host ("  전표{0} 부가세 결제카드: {1}카드인데 {2} 넣음 -> {3} 여야 함" -f $sq, $cardNm, $ctnb, $want) -ForegroundColor Red }
  }
}
$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

''
"검사한 줄: $checked 개"
if ($bad -eq 0) { Write-Host "틀린 곳 없습니다." -ForegroundColor Green }
else { Write-Host "틀린 곳 $bad 군데 -> 위를 보고 고쳐 주세요." -ForegroundColor Red }
