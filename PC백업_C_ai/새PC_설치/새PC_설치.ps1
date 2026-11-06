# 새 PC 설치 (처음 한 번만) — 이 폴더(새PC_설치)를 새 PC 아무 곳에 두고 실행
#  1) 엑셀 확인  2) Git 설치(없으면)  3) 회사 저장소 받기(C:\ai_github\novix-finance-yunheeseung)
#  4) 저장소 → C:\ai (도구·CLAUDE.md·클로드 메모리·기초자료·참고자료)  5) 기초자료용 date 저장소 받기  6) 10분 자동 동기화 등록
#  설치 뒤에는 노트북 ↔ 데스크탑이 10분마다 자동으로 맞춰짐 (더 새 파일이 이김, 지우지는 않음)
$ErrorActionPreference = 'Stop'
$repoUrl2 = "https://github.com/daeshin-novix/novix-finance-yunheeseung"; $repo2 = "C:\ai_github\novix-finance-yunheeseung"
$repoUrl = "https://github.com/a01042102350/date"; $repo = "C:\ai_github\date"

"1) 엑셀 확인"
if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application')) { throw "이 PC에 엑셀이 없습니다. 엑셀을 먼저 설치해 주세요." }
"   엑셀 있음"

"2) Git 확인"
$git = "C:\Program Files\Git\cmd\git.exe"
if (-not (Test-Path $git)) { "   Git 설치 중..."; winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements --silent; if (-not (Test-Path $git)) { throw "Git 설치 실패 — https://git-scm.com 에서 직접 설치해 주세요." } }
"   Git 있음"

"3) 회사 저장소 받기 (로그인 창이 뜨면 GitHub a01042102350 계정으로 로그인)"
New-Item -ItemType Directory -Force "C:\ai_github" | Out-Null
if (-not (Test-Path "$repo2\.git")) { & $git -c credential.helper=manager clone $repoUrl2 $repo2 } else { & $git -C $repo2 -c credential.helper=manager pull --rebase origin main }
if (-not (Test-Path "$repo2\.git")) { throw "회사 저장소를 받지 못했습니다. 로그인·권한을 확인해 주세요." }
& $git -C $repo2 config user.name "a01042102350"; & $git -C $repo2 config user.email "rlrn1234@ds-al.com"; & $git -C $repo2 config core.quotepath false

"4) 저장소 → C:\ai 복사"
$o = @('/E', '/XO', '/XF', '~$*', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
New-Item -ItemType Directory -Force "C:\ai\도구", "C:\ai\기초자료" | Out-Null
robocopy "$repo2\src\월마감" "C:\ai\도구" @o | Out-Null
Copy-Item "$repo2\docs\월마감_작업규칙.md" "C:\ai\CLAUDE.md" -Force
$mem = Join-Path $env:USERPROFILE ".claude\projects\c--ai\memory"; New-Item -ItemType Directory -Force $mem | Out-Null
if (Test-Path "$repo2\docs\클로드메모리") { robocopy "$repo2\docs\클로드메모리" $mem @o | Out-Null }
robocopy "$repo2\기초자료" "C:\ai\기초자료" @o | Out-Null
if (Test-Path "$repo2\참고자료") { robocopy "$repo2\참고자료" "C:\ai" @o | Out-Null }
"   C:\ai\도구, C:\ai\CLAUDE.md, 클로드 메모리($mem), C:\ai\기초자료, 참고자료 복사 완료"

"5) 기초자료용 date 저장소 받기"
if (-not (Test-Path "$repo\.git")) { & $git clone $repoUrl $repo } else { & $git -C $repo pull --ff-only origin main }
& $git -C $repo config user.name "a01042102350"; & $git -C $repo config core.quotepath false

"6) 10분 자동 동기화 등록"
$p = "C:\ai\도구\깃허브_자동동기화.ps1"
$act = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$p`""
$trg = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 10)
$set = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
Register-ScheduledTask -TaskName "월마감_깃허브_동기화" -Action $act -Trigger $trg -Settings $set -Description "C:\ai ↔ GitHub(회사 저장소) 10분마다 동기화" -Force | Out-Null
"   등록 완료"
""
"설치 끝. 이제 이 PC와 다른 PC의 C:\ai 가 10분마다 자동으로 맞춰집니다."
"VS Code에서 C:\ai 폴더를 열고 클로드에게 말하면 됩니다. 예) 202610 반영해줘"
