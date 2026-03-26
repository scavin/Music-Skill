---
name: go-music-api
description: Search songs, download playable audio, fetch lyrics, parse music share links, configure platform cookies, and switch music sources through a local go-music-api backend on Linux, macOS, and Windows. Use when the user asks to search music, play or download a song, get lyrics, resolve music share links, configure cookies for login-required tracks, or retry playback from another source after failures. If the backend is missing or unhealthy, install and start go-music-api first.
repository: https://github.com/scavin/Music-Skill
---

# go-music-api

Use this skill to install and run a local `go-music-api` backend on Linux, macOS, and Windows, search tracks across sources, download audio, embed metadata, and recover from source failures.

## Primary workflow

Prefer the bundled scripts instead of reimplementing the flow by hand.

### Install and start the backend

Run (Linux/macOS):

```bash
scripts/install.sh
```

Run (Windows PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/install.ps1
```

The install script should:
- install `go-music-api` into `~/.openclaw/music`
- choose a usable local port
- start the backend in the background
- verify health with a local API request

### Configure cookies when users mention login-only tracks

Trigger this flow when the user mentions cookie, VIP-only tracks, no-copyright/grey tracks, or asks to configure NetEase/QQ/Soda account state.

Prefer runtime API update instead of manual file placement:

Linux/macOS:

```bash
PORT="$(cat "$HOME/.openclaw/music/port" 2>/dev/null || echo 8080)"
curl -fsS -X POST "http://localhost:${PORT}/api/v1/system/cookies" \
  -H "Content-Type: application/json" \
  -d '{
    "netease": "MUSIC_U=xxx; __csrf=yyy;",
    "qq": "qm_keyst=xxx; uin=yyy;",
    "soda": "sessionid=xxx;"
  }'
curl -fsS "http://localhost:${PORT}/api/v1/system/cookies"
```

Windows PowerShell:

```powershell
$portFile = Join-Path $env:USERPROFILE ".openclaw\music\port"
$port = if (Test-Path $portFile) { (Get-Content $portFile).Trim() } else { "8080" }
$body = @{
  netease = "MUSIC_U=xxx; __csrf=yyy;"
  qq      = "qm_keyst=xxx; uin=yyy;"
  soda    = "sessionid=xxx;"
} | ConvertTo-Json
Invoke-RestMethod -Method Post -Uri "http://localhost:$port/api/v1/system/cookies" -ContentType "application/json" -Body $body
Invoke-RestMethod -Method Get  -Uri "http://localhost:$port/api/v1/system/cookies"
```

Notes:
- Do not print full cookie values back to the user unless explicitly requested.
- After setting cookies, retry search/stream for the same track before switching source.
- Upstream also supports startup loading from `cookies.json`; runtime API update is preferred in this skill.

### Download a song

Run (Linux/macOS):

```bash
scripts/play.sh "稻香" "$HOME/.openclaw/media/daoxiang.mp3"
```

Run (Windows PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/play.ps1 "稻香" "$env:USERPROFILE\.openclaw\media\daoxiang.mp3"
```

The play script should:
- search by query
- handle nested search payloads such as `data.data`, `data.list`, or `data.songs`
- rank candidates by song title, artist, and source quality
- avoid karaoke, cover, remix, live, DJ, and instrumental variants when possible
- download the audio stream to the requested path
- reuse cached files when an equivalent file already exists
- call `scripts/embed_metadata.py` to write title, artist, album, cover art, and embedded lyrics when available

Prefer saving final media under a sendable location such as `~/.openclaw/media/`.

## Manual API workflow

Use this only for debugging or when the helper scripts need changes.

1. Ensure the backend is installed and running with `scripts/install.sh`.
2. Read `~/.openclaw/music/port`; default to `8080` if absent.
3. Search with `GET /api/v1/music/search?q={q}`.
4. Parse list results from the top-level response or nested `data.*` collections.
5. Choose the best candidate.
6. Download audio with `GET /api/v1/music/stream?id={id}&source={source}`.
7. Treat the `stream` response as audio bytes, not JSON.
8. If playback fails, try `GET /api/v1/music/switch?...` to switch source and retry.
9. If user provides account cookies, set them with `POST /api/v1/system/cookies` and verify with `GET /api/v1/system/cookies`.
10. Fetch lyrics with `GET /api/v1/music/lyric?id={id}&source={source}` when needed.

## Files and state

Runtime files live under `~/.openclaw/music` (Linux/macOS) or `%USERPROFILE%\.openclaw\music` (Windows):
- binary: `go-music-api` (Linux/macOS) or `go-music-api.exe` (Windows)
- log: `log.txt`
- pid: `pid`
- port: `port`
- cache index: `cache-index.json`

## Failure handling

- If installation fails, check platform and architecture detection, GitHub Releases reachability, and required tools such as `curl`, `tar`, `unzip`, and `file` (Linux/macOS) or PowerShell 5.1+ (Windows).
- Match release asset names exactly. Do not use loose matching that could select `.deb` or `.rpm` packages.
- Accept only native executables after extraction. Fail immediately for text files, HTML, scripts, or package files. On Windows, validate the PE header (MZ signature).
- Windows only supports amd64 architecture (the only release asset available).
- If metadata embedding is required, ensure Python and `mutagen` are available. If not, skip metadata embedding or install the dependency before retrying.
- If certain tracks fail due to platform restrictions, ask for platform cookies and apply them via `/api/v1/system/cookies` before switching sources.
- If startup health checks fail, inspect `~/.openclaw/music/log.txt`.
- If the backend always binds to a fixed port in practice, simplify the port logic instead of pretending dynamic ports work.
