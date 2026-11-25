# 원챔버 '문서 더블클릭 = 읽기 모드'(서버 설정, 2026-11-24부터)를 피해서 엑셀 파일을 수정 모드로 여는 도구 (2026-11-25)
# 1) 오른쪽 클릭 → 더 많은 옵션 표시 → '엑셀 수정 모드로 열기'
# 2) 바탕화면 '엑셀 수정모드로 열기' 아이콘: 더블클릭하면 파일 고르는 창, 또는 엑셀 파일을 아이콘 위로 끌어다 놓기
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Files)
if (-not $Files -or -not $Files.Count) {
  Add-Type -AssemblyName System.Windows.Forms
  $d = New-Object System.Windows.Forms.OpenFileDialog
  $d.Title = '수정 모드로 열 엑셀 파일을 고르세요'; $d.InitialDirectory = 'C:\ai\기초자료'; $d.Filter = '엑셀 파일 (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|모든 파일 (*.*)|*.*'; $d.Multiselect = $true
  if ($d.ShowDialog() -ne 'OK') { return }
  $Files = $d.FileNames
}
foreach ($f in $Files) { if ($f -and (Test-Path -LiteralPath $f)) { Start-Process -FilePath $f } }
