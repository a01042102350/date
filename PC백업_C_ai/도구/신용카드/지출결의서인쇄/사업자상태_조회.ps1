param(
  [string]$Data = "$env:TEMP\카드전표\자료.json",
  [string]$Out = "$env:TEMP\카드전표\사업자상태.csv"
)
# 영수증의 가맹점 사업자번호 → 국세청(홈택스) 사업자등록 상태조회(로그인 없는 공개 조회) → 일반/간이/면세/폐업
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web.Extensions
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer; $ser.MaxJsonLength = [int]::MaxValue
$docs = $ser.DeserializeObject([IO.File]::ReadAllText($Data, [Text.Encoding]::UTF8))
$biz = [ordered]@{}
foreach ($d in $docs) { foreach ($it in $d['items']) { $rc = $it['receipt']; if ($rc) { $b = ("$($rc['bizNo'])" -replace '\D', ''); if ($b.Length -eq 10 -and -not $biz.Contains($b)) { $biz[$b] = "$($rc['name'])" } } } }
$res = New-Object Collections.ArrayList
foreach ($b in $biz.Keys) {
  $body = "{`"mobYn`":`"N`",`"inqrTrgtClCd`":`"1`",`"txprDscmNo`":`"$b`",`"dongCode`":`"__`",`"psbSearch`":`"Y`"}"
  $txt = ''; $kind = '조회 실패'
  try {
    $r = Invoke-WebRequest -Uri 'https://teht.hometax.go.kr/wqAction.do?actionId=ATTABZAA001R08&screenId=UTEABAAA13&popupYn=false&realScreenId=' -Method Post -Body ([Text.Encoding]::UTF8.GetBytes($body)) -ContentType 'application/json; charset=UTF-8' -UseBasicParsing -TimeoutSec 20
    $o = $ser.DeserializeObject([Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()))
    $txt = ("$($o['trtCntn'])" -replace '\s+', ' ').Trim()
    $kind = if ($txt -match '폐업') { '폐업' } elseif ($txt -match '휴업') { '휴업' } elseif ($txt -match '간이과세자\(세금계산서 발급사업자\)') { '간이(세금계산서 발급)' } elseif ($txt -match '간이과세자') { '간이' } elseif ($txt -match '면세') { '면세' } elseif ($txt -match '일반과세자') { '일반' } elseif ($txt -match '비영리|국가|지방자치|고유번호') { '비영리·기관' } elseif ($txt -match '등록되어 있지 않') { '미등록' } else { '기타' }
  } catch { $txt = $_.Exception.Message }
  [void]$res.Add([pscustomobject]@{ 사업자번호 = $b; 가맹점 = $biz[$b]; 구분 = $kind; 국세청결과 = $txt })
  "{0} {1}: {2}" -f $b, $biz[$b], $kind
  Start-Sleep -Milliseconds 600
}
$res | Export-Csv $Out -NoTypeInformation -Encoding UTF8
"저장: $Out ($($res.Count)곳)"
