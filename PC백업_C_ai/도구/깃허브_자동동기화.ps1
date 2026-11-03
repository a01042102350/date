# C:\ai ↔ GitHub 자동 동기화 — 작업 스케줄러가 10분마다 실행
#  C:\ai\기초자료 ↔ GitHub 기초자료/,  C:\ai의 나머지 전부 ↔ GitHub PC백업_C_ai/ (2026-11-03 추가)
#  2) GitHub → PC: 다른 곳에서 GitHub에 올린 파일·폴더를 이 PC 폴더로 (더 새 파일만)
#  3) PC → GitHub: 이 PC 폴더에 넣은 파일·폴더를 GitHub로 (더 새 파일만, 지우지는 않음)
#  빈 폴더도 GitHub에 남도록 .gitkeep 를 넣어 둠 (PC 쪽으로는 안 가져옴)
$git = "C:\Program Files\Git\cmd\git.exe"; $repo = "C:\ai_github\date"; $local = "C:\ai\기초자료"; $url = "https://github.com/a01042102350/date"
$log = "C:\ai\도구\동기화기록.txt"
function Log($m) { Add-Content -Path $log -Value ("{0:yyyy-MM-dd HH:mm}  {1}" -f (Get-Date), $m) -Encoding UTF8 }

# 공개 저장소여도 올림 (사용자 결정 2026-10-29: "공개여도 상관없어")
$env:GIT_TERMINAL_PROMPT = '1'

& $git -C $repo pull --ff-only origin main 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { Log "실패: GitHub에서 받아오기(pull) 안 됨 → 클로드에게 확인 요청"; exit }
$opt = @('/E', '/XO', '/XF', '~$*', 'Thumbs.db', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
robocopy "$repo\기초자료" $local @opt | Out-Null     # GitHub → PC
robocopy $local "$repo\기초자료" @opt | Out-Null     # PC → GitHub

# C:\ai 나머지 (기초자료는 위에서, .git·동기화기록은 제외)
$all = "C:\ai"; $bak = "$repo\PC백업_C_ai"
$opt2 = @('/E', '/XO', '/XD', "$all\.git", "$all\기초자료", "$bak\.git", "$bak\기초자료", '/XF', '~$*', 'Thumbs.db', '.gitkeep', '동기화기록.txt', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
robocopy $bak $all @opt2 | Out-Null     # GitHub → PC
robocopy $all $bak @opt2 | Out-Null     # PC → GitHub

# 빈 폴더 표시용 .gitkeep
foreach ($top in "$repo\기초자료", $bak) {
  Get-ChildItem $top -Directory -Recurse -Force | Where-Object { -not (Get-ChildItem $_.FullName -Force) } | ForEach-Object { New-Item -ItemType File -Path (Join-Path $_.FullName '.gitkeep') | Out-Null }
}

& $git -C $repo add -A "기초자료" "PC백업_C_ai" 2>&1 | Out-Null
& $git -C $repo diff --cached --quiet
if ($LASTEXITCODE -eq 0) { exit }   # 바뀐 것 없음
$files = @(& $git -C $repo -c core.quotepath=false diff --cached --name-only)
& $git -C $repo commit -m ("자동 동기화 {0:yyyy-MM-dd HH:mm} ({1}개 파일)" -f (Get-Date), $files.Count) 2>&1 | Out-Null
& $git -C $repo -c credential.helper=manager push origin main 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) { Log ("올림 {0}개: {1}" -f $files.Count, (($files | select -First 5) -join ', ')) } else { Log "실패: GitHub에 올리기(push) 안 됨 → 로그인 확인 필요" }
