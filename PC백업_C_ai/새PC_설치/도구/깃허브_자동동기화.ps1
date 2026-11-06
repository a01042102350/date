# C:\ai ↔ GitHub 자동 동기화 — 작업 스케줄러가 10분마다 실행 (회사 저장소 = 노트북·데스크탑 연동, date 저장소 = 기초자료)
#  2) GitHub → PC: 다른 곳에서 GitHub에 올린 파일을 이 PC 폴더로 (더 새 파일만)
#  3) PC → GitHub: 이 PC 폴더에 넣은 파일을 GitHub로 (더 새 파일만, 지우지는 않음)
$git = "C:\Program Files\Git\cmd\git.exe"; $repo = "C:\ai_github\date"; $local = "C:\ai\기초자료"; $url = "https://github.com/a01042102350/date"
$log = "C:\ai\도구\동기화기록.txt"
function Log($m) { Add-Content -Path $log -Value ("{0:yyyy-MM-dd HH:mm}  {1}" -f (Get-Date), $m) -Encoding UTF8 }

# 공개 저장소여도 올림 (사용자 결정 2026-10-29: "공개여도 상관없어")
$env:GIT_TERMINAL_PROMPT = '1'

# 회사 저장소(novix-finance-yunheeseung, 비공개): 엑셀·PDF·HTML까지 C:\ai 자료 전부 올림 — 어디서든 보게 (사용자 결정 2026-11-03)
#   노트북 ↔ 데스크탑 연동 (사용자 요청 2026-11-03): 모든 항목 양쪽으로, 더 새 파일이 이김, 지우지는 않음
#   C:\ai\도구 ↔ src\월마감\ / CLAUDE.md ↔ docs\월마감_작업규칙.md / 클로드 메모리 ↔ docs\클로드메모리\
#   C:\ai\기초자료 ↔ 기초자료\ / C:\ai 나머지 폴더·파일 ↔ 참고자료\ (.zip은 저장소 규칙상 제외, 동기화기록.txt는 PC마다 따로)
$repo2 = "C:\ai_github\novix-finance-yunheeseung"
$opt2 = @('/E', '/XO', '/XF', '~$*', 'Thumbs.db', '*.zip', '동기화기록.txt', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
function SyncDir($a, $b) { New-Item -ItemType Directory -Force $a, $b | Out-Null; robocopy $b $a @opt2 | Out-Null; robocopy $a $b @opt2 | Out-Null }
function SyncFile($a, $b) {   # 이름이 다른 파일 한 쌍: 더 새 쪽으로 맞춤 (2초 이내 차이는 같은 것으로 봄)
  $fa = Get-Item $a -ErrorAction SilentlyContinue; $fb = Get-Item $b -ErrorAction SilentlyContinue
  if (-not $fa -and -not $fb) { return }
  if (-not $fb -or ($fa -and ($fa.LastWriteTimeUtc - $fb.LastWriteTimeUtc).TotalSeconds -gt 2)) { New-Item -ItemType Directory -Force (Split-Path $b) | Out-Null; Copy-Item $a $b -Force }
  elseif (-not $fa -or ($fb.LastWriteTimeUtc - $fa.LastWriteTimeUtc).TotalSeconds -gt 2) { New-Item -ItemType Directory -Force (Split-Path $a) | Out-Null; Copy-Item $b $a -Force }
}
if (Test-Path "$repo2\.git") {
  & $git -C $repo2 -c credential.helper=manager pull --rebase origin main 2>&1 | Out-Null
  if ($LASTEXITCODE -ne 0) { & $git -C $repo2 rebase --abort 2>&1 | Out-Null; Log "실패: 회사 저장소에서 받아오기(pull) 안 됨 → 클로드에게 확인 요청" }
  else {
    SyncDir "C:\ai\도구" "$repo2\src\월마감"
    SyncFile "C:\ai\CLAUDE.md" "$repo2\docs\월마감_작업규칙.md"
    SyncDir (Join-Path $env:USERPROFILE ".claude\projects\c--ai\memory") "$repo2\docs\클로드메모리"
    SyncFile "$local\기초자료_넣는순서.txt" "$repo2\docs\기초자료_넣는순서.txt"
    SyncDir $local "$repo2\기초자료"
    $names = @(@(Get-ChildItem "C:\ai" -Directory) + @(Get-ChildItem "$repo2\참고자료" -Directory -ErrorAction SilentlyContinue) | % Name | ? { $_ -notin '기초자료', '도구' } | select -Unique)
    foreach ($nm in $names) { SyncDir "C:\ai\$nm" "$repo2\참고자료\$nm" }
    $fn = @(@(Get-ChildItem "C:\ai" -File) + @(Get-ChildItem "$repo2\참고자료" -File -ErrorAction SilentlyContinue) | % Name | ? { $_ -notlike '*.zip' -and $_ -notlike '~$*' -and $_ -ne 'CLAUDE.md' } | select -Unique)
    foreach ($nm in $fn) { SyncFile "C:\ai\$nm" "$repo2\참고자료\$nm" }
    & $git -C $repo2 add -A "src/월마감" "docs" "기초자료" "참고자료" 2>&1 | Out-Null
    & $git -C $repo2 diff --cached --quiet
    if ($LASTEXITCODE -ne 0) {
      $files2 = @(& $git -C $repo2 -c core.quotepath=false diff --cached --name-only)
      & $git -C $repo2 commit -m ("월마감 자료 동기화 {0:yyyy-MM-dd HH:mm} ({1}개 파일)" -f (Get-Date), $files2.Count) 2>&1 | Out-Null
      & $git -C $repo2 -c credential.helper=manager push origin main 2>&1 | Out-Null
      if ($LASTEXITCODE -eq 0) { Log ("회사 저장소 올림 {0}개: {1}" -f $files2.Count, (($files2 | select -First 5) -join ', ')) } else { Log "실패: 회사 저장소에 올리기(push) 안 됨 → 로그인·권한 확인 필요" }
    }
  }
}

& $git -C $repo pull --ff-only origin main 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { Log "실패: GitHub에서 받아오기(pull) 안 됨 → 클로드에게 확인 요청"; exit }
$opt = @('/E', '/XO', '/XF', '~$*', 'Thumbs.db', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
robocopy "$repo\기초자료" $local @opt | Out-Null     # GitHub → PC
robocopy $local "$repo\기초자료" @opt | Out-Null     # PC → GitHub

# C:\ai 나머지 전부 ↔ PC백업_C_ai/ — 노트북 ↔ 데스크탑 연동 (데스크탑 방식 합침 2026-11-06). 기초자료·.git·동기화기록은 제외
$all = "C:\ai"; $bak = "$repo\PC백업_C_ai"
$opt3 = @('/E', '/XO', '/XD', "$all\.git", "$all\기초자료", "$bak\.git", "$bak\기초자료", '/XF', '~$*', 'Thumbs.db', '.gitkeep', '동기화기록.txt', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
robocopy $bak $all @opt3 | Out-Null     # GitHub → PC
robocopy $all $bak @opt3 | Out-Null     # PC → GitHub
# 빈 폴더 표시용 .gitkeep (PC 쪽으로는 안 가져옴)
foreach ($top in "$repo\기초자료", $bak) {
  Get-ChildItem $top -Directory -Recurse -Force | ? { -not (Get-ChildItem $_.FullName -Force) } | % { New-Item -ItemType File -Path (Join-Path $_.FullName '.gitkeep') | Out-Null }
}

& $git -C $repo add -A "기초자료" "PC백업_C_ai" 2>&1 | Out-Null
& $git -C $repo diff --cached --quiet
if ($LASTEXITCODE -eq 0) { exit }   # 바뀐 것 없음
$files = @(& $git -C $repo -c core.quotepath=false diff --cached --name-only)
& $git -C $repo commit -m ("자동 동기화 {0:yyyy-MM-dd HH:mm} ({1}개 파일)" -f (Get-Date), $files.Count) 2>&1 | Out-Null
& $git -C $repo -c credential.helper=manager push origin main 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) { Log ("올림 {0}개: {1}" -f $files.Count, (($files | select -First 5) -join ', ')) } else { Log "실패: GitHub에 올리기(push) 안 됨 → 로그인 확인 필요" }
