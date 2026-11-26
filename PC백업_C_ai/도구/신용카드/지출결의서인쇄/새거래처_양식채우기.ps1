# 아마란스 거래처엑셀업로드 양식(양식만들기로 받은 파일)에 새 거래처를 채워 넣음
#  - 엑셀로 다시 저장하지 않고 양식 파일 그대로에 줄만 더함 (아마란스가 만든 모양 그대로)
#  - 양식 2줄의 칸코드(CO_CD, TR_FG, TR_CD ...)를 읽어, 양식에 있는 칸만 채움 → 양식을 바꿔도 그대로 씀
#  - Records: 칸코드를 키로 하는 표(hashtable) 목록. 예) @{ CO_CD='1000'; TR_FG='1'; TR_CD='05724'; TR_NM='진곰탕'; ATTR_NM='진곰탕'; REG_NB='6103391121' }
param([Parameter(Mandatory)][string]$Template, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][object[]]$Records)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function ColName([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][math]::Floor(($n - 1) / 26) }; $s }
function ColNum([string]$c) { $n = 0; foreach ($ch in $c.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; $n }
function Read-Entry($zip, $name) { $e = $zip.GetEntry($name); if (-not $e) { return $null }; $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); try { $sr.ReadToEnd() } finally { $sr.Close() } }
function Write-Entry($zip, $name, $text) { $old = $zip.GetEntry($name); if ($old) { $old.Delete() }; $ne = $zip.CreateEntry($name); $sw = New-Object IO.StreamWriter($ne.Open(), (New-Object Text.UTF8Encoding($false))); try { $sw.Write($text) } finally { $sw.Close() } }
function Esc([string]$s) { [Security.SecurityElement]::Escape($s) }

if (-not (Test-Path -LiteralPath $Template)) { throw "거래처 업로드 양식이 없습니다: $Template" }
$tmp = "$Out.임시"; if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
Copy-Item -LiteralPath $Template $tmp -Force
$z = [IO.Compression.ZipFile]::Open($tmp, 'Update')
try {
  $sheetPath = 'xl/worksheets/sheet1.xml'
  $sheet = Read-Entry $z $sheetPath; $sst = Read-Entry $z 'xl/sharedStrings.xml'
  if (-not $sheet -or -not $sst) { throw '양식 파일 모양이 예상과 다릅니다(sheet1 또는 공유 문자열 없음)' }
  [xml]$ssx = $sst
  $strs = @($ssx.sst.si | ForEach-Object { if ($_.t -is [string]) { $_.t } elseif ($_.t) { $_.t.'#text' } else { ($_.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '' } })

  # 2줄 칸코드 → 열 번호
  $row2 = [regex]::Match($sheet, '(?s)<row [^>]*r="2"[^>]*>(.*?)</row>').Groups[1].Value
  $col = [ordered]@{}
  foreach ($cm in [regex]::Matches($row2, '(?s)<c r="([A-Z]+)2"([^>]*?)(?:/>|>(.*?)</c>)')) {
    $v = [regex]::Match($cm.Groups[3].Value, '<v>(.*?)</v>').Groups[1].Value
    if ($cm.Groups[2].Value -match 't="s"' -and $v -ne '') { $v = $strs[[int]$v] }
    $v = "$v".Trim(); if ($v -and -not $col.Contains($v)) { $col[$v] = ColNum $cm.Groups[1].Value }
  }
  foreach ($need in 'CO_CD', 'TR_FG', 'TR_CD', 'TR_NM', 'ATTR_NM') { if (-not $col.Contains($need)) { throw "거래처 양식에 꼭 있어야 할 칸이 없음: $need" } }
  if ($sheet -match '<row [^>]*r="4"') { throw '양식 4줄에 이미 자료가 있습니다 → 빈 양식(양식만들기로 받은 그대로)을 써 주세요' }

  $uniq = [int][regex]::Match($sst, 'uniqueCount="(\d+)"').Groups[1].Value
  $cnt = [int][regex]::Match($sst, '\scount="(\d+)"').Groups[1].Value
  $idx = @{}; for ($i = 0; $i -lt $strs.Count; $i++) { if (-not $idx.ContainsKey($strs[$i])) { $idx[$strs[$i]] = $i } }
  $addSi = New-Object Text.StringBuilder
  $rowsXml = New-Object Text.StringBuilder
  $byNum = @($col.Keys | Sort-Object { $col[$_] })
  $rn = 4
  foreach ($rec in $Records) {
    [void]$rowsXml.Append("<row r=`"$rn`">")
    foreach ($k in $byNum) {
      $val = if ($rec -is [System.Collections.IDictionary]) { $rec[$k] } else { $rec.$k }
      $val = "$val".Trim(); if (-not $val) { continue }
      if (-not $idx.ContainsKey($val)) { $idx[$val] = $uniq; $uniq++; [void]$addSi.Append('<si><t xml:space="preserve">' + (Esc $val) + '</t></si>') }
      $cnt++
      $st = if ($col[$k] -le 5) { 4 } else { 5 }   # 양식 열 기본 서식(문자) 그대로
      [void]$rowsXml.Append("<c r=`"$(ColName $col[$k])$rn`" s=`"$st`" t=`"s`"><v>$($idx[$val])</v></c>")
    }
    [void]$rowsXml.Append('</row>'); $rn++
  }
  $lastRow = $rn - 1
  $lastCol = ColName (($col.Values | Measure-Object -Maximum).Maximum)
  $sheet = $sheet.Replace('</sheetData>', $rowsXml.ToString() + '</sheetData>')
  $sheet = [regex]::Replace($sheet, '<dimension ref="[^"]*"/>', "<dimension ref=`"A1:$lastCol$lastRow`"/>")
  $sst = $sst.Replace('</sst>', $addSi.ToString() + '</sst>')
  $sst = [regex]::Replace($sst, 'uniqueCount="\d+"', "uniqueCount=`"$uniq`"")
  $sst = [regex]::Replace($sst, '(\s)count="\d+"', "`${1}count=`"$cnt`"")
  [void][xml]$sheet; [void][xml]$sst   # 깨진 XML이면 여기서 멈춤
  Write-Entry $z $sheetPath $sheet
  Write-Entry $z 'xl/sharedStrings.xml' $sst
} finally { $z.Dispose() }
if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
Move-Item -LiteralPath $tmp $Out -Force
"새 거래처 업로드 파일: $Out ($($Records.Count)곳, 칸: $($byNum -join ', '))"
