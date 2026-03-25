$ErrorActionPreference = "Stop"

$Query    = $args[0]
$OutPath  = $args[1]
$PortFile = Join-Path $env:USERPROFILE ".openclaw\music\port"
$CacheIdx = Join-Path $env:USERPROFILE ".openclaw\music\cache-index.json"
$DefaultDir = Join-Path $env:USERPROFILE ".openclaw\media"

if (-not $Query) {
    Write-Host "usage: play.ps1 <query> [output-path]" -ForegroundColor Red
    exit 1
}

# Locate python: py launcher > python (verify 3.x) > python3
$pyExe = $null
$pyArgs = @()
# 1. py -3 (Windows Python Launcher, most reliable)
try {
    $ver = & py -3 --version 2>&1
    if ($ver -match "Python 3") { $pyExe = "py"; $pyArgs = @("-3") }
} catch {}
# 2. python (skip if it's the Microsoft Store stub)
if (-not $pyExe) {
    try {
        $pyPath = (Get-Command python -ErrorAction Stop).Source
        if ($pyPath -notmatch "WindowsApps") {
            $ver = & python --version 2>&1
            if ($ver -match "Python 3") { $pyExe = "python" }
        }
    } catch {}
}
# 3. python3
if (-not $pyExe) {
    try {
        $ver = & python3 --version 2>&1
        if ($ver -match "Python 3") { $pyExe = "python3" }
    } catch {}
}
# 4. Auto-install via winget
if (-not $pyExe) {
    Write-Host "[music] python not found, attempting install via winget..."
    try {
        & winget install --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null
        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        $ver = & py -3 --version 2>&1
        if ($ver -match "Python 3") { $pyExe = "py"; $pyArgs = @("-3") }
    } catch {}
}
if (-not $pyExe) {
    Write-Host "[music] python 3 not found and auto-install failed. Please install from https://www.python.org/downloads/" -ForegroundColor Red
    exit 1
}

if (-not $OutPath) {
    $safeName = ($Query -replace '[/ ]', '__' -replace '[^\w.\-\u4e00-\u9fff]', '')
    if (-not $safeName) { $safeName = "music" }
    $OutPath = Join-Path $DefaultDir "$safeName.mp3"
}

$outDir = Split-Path $OutPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$cacheDir = Split-Path $CacheIdx -Parent
if (-not (Test-Path $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }

if (Test-Path $PortFile) {
    $Port = (Get-Content $PortFile -ErrorAction SilentlyContinue).Trim()
}
if (-not $Port) { $Port = "8080" }

# --- Cache check ---
$cacheCheckPy = @"
import json, os, sys, shutil
query, out_path, cache_index = sys.argv[1:4]
if os.path.exists(out_path) and os.path.getsize(out_path) > 1024:
    print(json.dumps({'cache_hit': True, 'path': out_path, 'size': os.path.getsize(out_path), 'reason': 'output_exists'}, ensure_ascii=False))
    raise SystemExit(0)
if os.path.exists(cache_index):
    try:
        idx = json.load(open(cache_index, 'r', encoding='utf-8'))
    except Exception:
        idx = {}
    hit = idx.get(query)
    if isinstance(hit, dict):
        p = hit.get('path')
        if p and os.path.exists(p) and os.path.getsize(p) > 1024:
            if p != out_path:
                os.makedirs(os.path.dirname(out_path), exist_ok=True)
                shutil.copy2(p, out_path)
            print(json.dumps({'cache_hit': True, 'path': out_path, 'size': os.path.getsize(out_path), 'reason': 'query_index', 'chosen': hit.get('chosen')}, ensure_ascii=False))
            raise SystemExit(0)
print('{}')
"@

$cacheResult = $cacheCheckPy | & $pyExe @pyArgs - $Query $OutPath $CacheIdx
if ($cacheResult -match '"cache_hit":\s*true') {
    Write-Host $cacheResult
    exit 0
}

# --- Search ---
$searchUrl = "http://localhost:$Port/api/v1/music/search?q=$([uri]::EscapeDataString($Query))"
$searchJson = (Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -ErrorAction Stop).Content

# --- Score, select, download ---
$downloadPy = @"
import json, os, re, shutil, sys, urllib.parse, urllib.request
raw, port, query, out_path, cache_index = sys.argv[1:6]
data = json.loads(raw)
items = []
if isinstance(data, dict):
    if isinstance(data.get('data'), list):
        items = data['data']
    elif isinstance(data.get('data'), dict):
        inner = data['data']
        if isinstance(inner.get('data'), list):
            items = inner['data']
        elif isinstance(inner.get('list'), list):
            items = inner['list']
        elif isinstance(inner.get('songs'), list):
            items = inner['songs']
    elif isinstance(data.get('list'), list):
        items = data['list']
if not items:
    raise SystemExit('no items parsed from search response')

query_l = query.lower()
prefer_artist = None
for name in ('\u5468\u6770\u4f26', 'jay chou', 'jay', '\u6797\u4fca\u6770', '\u9648\u5955\u8fc5'):
    if name in query_l:
        prefer_artist = name
        break

def slug(text):
    text = re.sub(r'[^\w\-\u4e00-\u9fff]+', '_', text, flags=re.UNICODE)
    text = re.sub(r'_+', '_', text).strip('_')
    return text or 'music'

def score(item):
    name = str(item.get('name', ''))
    artist = str(item.get('artist', ''))
    source = str(item.get('source', ''))
    s = 0
    for token in query.replace('/', ' ').split():
        if token and token in name:
            s += 80
        if token and token in artist:
            s += 50
    if prefer_artist and prefer_artist.replace(' ', '').lower() in artist.replace(' ', '').lower():
        s += 120
    if source in ('migu', 'qq', 'netease', 'kuwo', 'kugou'):
        s += 20
    bad_words = ('\u4f34\u594f', '\u5c0f\u63d0\u7434', '\u7ffb\u5531', 'cover', 'live', 'dj', 'remix', '\u7eaf\u97f3\u4e50')
    if any(word.lower() in (name + ' ' + artist).lower() for word in bad_words):
        s -= 100
    return s

chosen = sorted(items, key=score, reverse=True)[0]
default_name = f"{slug(chosen.get('artist','unknown'))}-{slug(chosen.get('name','music'))}.mp3"
default_path = os.path.join(os.path.expanduser('~/.openclaw/media'), default_name)

if os.path.exists(default_path) and os.path.getsize(default_path) > 1024:
    if os.path.abspath(default_path) != os.path.abspath(out_path):
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        shutil.copy2(default_path, out_path)
    try:
        idx = json.load(open(cache_index, 'r', encoding='utf-8')) if os.path.exists(cache_index) else {}
    except Exception:
        idx = {}
    idx[query] = {'path': out_path, 'chosen': chosen}
    json.dump(idx, open(cache_index, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(json.dumps({'cache_hit': True, 'path': out_path, 'size': os.path.getsize(out_path), 'reason': 'song_file_exists', 'chosen': chosen}, ensure_ascii=False))
    raise SystemExit(0)

os.makedirs(os.path.dirname(out_path), exist_ok=True)
params = urllib.parse.urlencode({'id': chosen['id'], 'source': chosen['source']})
url = f'http://localhost:{port}/api/v1/music/stream?{params}'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
with urllib.request.urlopen(req) as resp, open(out_path, 'wb') as f:
    while True:
        chunk = resp.read(65536)
        if not chunk:
            break
        f.write(chunk)
if os.path.abspath(default_path) != os.path.abspath(out_path):
    os.makedirs(os.path.dirname(default_path), exist_ok=True)
    shutil.copy2(out_path, default_path)
try:
    idx = json.load(open(cache_index, 'r', encoding='utf-8')) if os.path.exists(cache_index) else {}
except Exception:
    idx = {}
idx[query] = {'path': out_path, 'chosen': chosen}
json.dump(idx, open(cache_index, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
print(json.dumps({'cache_hit': False, 'chosen': chosen, 'path': out_path, 'canonical_path': default_path, 'size': os.path.getsize(out_path)}, ensure_ascii=False))
"@

$downloadPy | & $pyExe @pyArgs - $searchJson $Port $Query $OutPath $CacheIdx

# --- Extract metadata for embedding ---
$metaExtractPy = @"
import json, os, sys, urllib.parse
query, out_path, cache_index, port = sys.argv[1:5]
idx = {}
if os.path.exists(cache_index):
    try:
        idx = json.load(open(cache_index, 'r', encoding='utf-8'))
    except Exception:
        idx = {}
entry = idx.get(query, {})
chosen = entry.get('chosen') or {}
song_json = json.dumps(chosen, ensure_ascii=False)
cover_url = chosen.get('cover', '')
lyric_url = ''
id_ = chosen.get('id', '')
source = chosen.get('source', '')
if id_ and source:
    lyric_url = f"http://localhost:{port}/api/v1/music/lyric?" + urllib.parse.urlencode({'id': id_, 'source': source})
print(f"{song_json}\n{cover_url}\n{lyric_url}")
"@

$metaOutput = $metaExtractPy | & $pyExe @pyArgs - $Query $OutPath $CacheIdx $Port
$metaLines = $metaOutput -split "`n"
$songJson  = $metaLines[0]
$coverUrl  = if ($metaLines.Length -gt 1) { $metaLines[1].Trim() } else { "" }
$lyricUrl  = if ($metaLines.Length -gt 2) { $metaLines[2].Trim() } else { "" }

# --- Embed metadata ---
$embedScript = Join-Path $PSScriptRoot "embed_metadata.py"
if (Test-Path $embedScript) {
    try {
        & $pyExe @pyArgs $embedScript $OutPath $Port $songJson $coverUrl $lyricUrl 2>&1 | Out-Null
    } catch {}
}
