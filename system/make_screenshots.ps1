# 화면 스크린샷 생성 — 헤드리스 Chrome
#
# 사용: 먼저 저장소 루트에서 정적 서버를 띄운다
#   python -m http.server 8820 --bind 127.0.0.1
# 그 다음
#   powershell -ExecutionPolicy Bypass -File system\make_screenshots.ps1
#
# AI 패널(15번)은 클릭이 필요해 별도 하네스가 있어야 한다 — 문서 참조.

$chrome = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$out    = "F:\policy_maps\docs\screenshots"
$base   = "http://127.0.0.1:8820/viz/public/index.html"
New-Item -ItemType Directory -Force -Path $out | Out-Null

$shots = @(
  @{n="01_dashboard";     h="#/dashboard";      ht=1400},
  @{n="02_map";           h="#/map";            ht=1200},
  @{n="03_region";        h="#/region/11110";   ht=1400},
  @{n="04_gap";           h="#/gap";            ht=1300},
  @{n="05_graph";         h="#/graph";          ht=1300},
  @{n="06_neural";        h="#/neural";         ht=1500},
  @{n="07_spatial";       h="#/spatial";        ht=1400},
  @{n="08_analytics";     h="#/analytics";      ht=1500},
  @{n="09_diffusion";     h="#/diffusion";      ht=1300},
  @{n="10_trust";         h="#/trust";          ht=1500},
  @{n="11_votes";         h="#/votes";          ht=1200},
  @{n="12_search";        h="#/search";         ht=1100},
  @{n="13_lifecycle";     h="#/lifecycle";      ht=1300},
  @{n="14_effectiveness"; h="#/effectiveness";  ht=1300}
)
foreach ($s in $shots) {
  $f = Join-Path $out "$($s.n).png"
  $a = @("--headless=new","--disable-gpu","--hide-scrollbars",
         "--window-size=1600,$($s.ht)","--virtual-time-budget=25000",
         "--screenshot=$f","$base$($s.h)")
  & $chrome @a 2>&1 | Out-Null
  if (Test-Path $f) { "  ok  $($s.n).png" } else { "  x   $($s.n)" }
}
