# Package Manager Distribution Strategy

**Summary:**
- QuickSheet has zero presence on any package manager. Current releases are tag-only (no binaries).
- Pre-built self-contained binaries on GitHub Releases are the prerequisite for ALL package managers.
- Highest ROI channels: **Scoop** (easy, dev audience), **winget** (wide Windows reach), **AUR** (Arch/unixporn crowd).
- A GitHub Actions CI workflow to publish binaries on tag is the single blocking step before any package manager submission.
- Once binaries exist, Scoop submission is < 1 hour of work; winget ~ 2 hours; AUR ~ 3 hours.

---

## Current State (2026-05-29)

- **Releases:** v0.36.0 is latest. All releases are source-only tags — zero binary assets attached.
- **CI:** No `.github/workflows/` directory exists. No automated build or release pipeline.
- **Install method:** Clone → `dotnet build` → run. This is a huge barrier for non-.NET users.

## Target Package Managers (ranked by ROI)

### 1. Scoop (Windows) — HIGH priority

| Attribute | Detail |
|-----------|--------|
| Audience | Windows developers, CLI power users |
| Format | JSON manifest in `ScoopInstaller/Extras` bucket |
| Requirements | Stable URL to a `.zip` of binaries, SHA256 hash |
| Effort | Low — one JSON file, test locally, submit PR |
| Overlap | r/commandline, r/dotnet, terminal-tool communities |

**Manifest sketch:**
```json
{
  "version": "0.36.0",
  "description": "Interactive spreadsheet as your desktop wallpaper",
  "homepage": "https://github.com/cemheren/QuickSheet",
  "license": "MIT",
  "architecture": {
    "64bit": {
      "url": "https://github.com/cemheren/QuickSheet/releases/download/v0.36.0/QuickSheet-win-x64.zip",
      "hash": "<sha256>"
    }
  },
  "bin": "QuickSheet.exe",
  "shortcuts": [["QuickSheet.exe", "QuickSheet"]]
}
```

### 2. winget (Windows) — HIGH priority

| Attribute | Detail |
|-----------|--------|
| Audience | All Windows 10/11 users (ships with OS) |
| Format | YAML manifest in `microsoft/winget-pkgs` |
| Requirements | Public installer URL (zip/exe/msi), SHA256, silent-install capability |
| Effort | Moderate — use `winget-create` tool, fill metadata, submit PR |
| Notes | Self-contained .NET publish zip works; "portable" installer type |

**Identifier:** `cemheren.QuickSheet`

### 3. AUR (Arch Linux) — MEDIUM-HIGH priority

| Attribute | Detail |
|-----------|--------|
| Audience | Arch Linux users, r/unixporn, tiling WM enthusiasts |
| Format | PKGBUILD script |
| Requirements | Source tarball or binary release URL |
| Effort | Moderate — write PKGBUILD, test in clean chroot |
| Notes | Can do `-bin` variant (pre-built) or source build. Binary is simpler for users. |

**PKGBUILD sketch (binary variant, `quicksheet-bin`):**
```bash
pkgname=quicksheet-bin
pkgver=0.36.0
pkgrel=1
pkgdesc="Interactive spreadsheet as your desktop wallpaper (X11)"
arch=('x86_64')
url="https://github.com/cemheren/QuickSheet"
license=('MIT')
depends=('libx11' 'libxft')
source=("$url/releases/download/v${pkgver}/QuickSheet-linux-x64.tar.gz")
sha256sums=('...')

package() {
  install -Dm755 QuickSheet "$pkgdir/usr/bin/quicksheet"
}
```

### 4. Homebrew (macOS/Linux) — MEDIUM priority

| Attribute | Detail |
|-----------|--------|
| Audience | macOS developers, some Linux users |
| Format | Ruby formula in a tap (`cemheren/homebrew-quicksheet`) |
| Requirements | Tarball URL, depends on .NET runtime or self-contained binary |
| Notes | QuickSheet doesn't support macOS desktop mode — console/headless only. Lower priority. |

### 5. Nix/nixpkgs — LOW priority (high effort)

| Attribute | Detail |
|-----------|--------|
| Audience | NixOS enthusiasts, declarative-config fans |
| Format | Nix derivation |
| Requirements | Complex — needs buildDotnetModule or fetchNuGet deps |
| Notes | Zero NuGet deps actually makes this EASIER than most .NET apps. Still high initial effort. |

### 6. Chocolatey (Windows) — LOW priority

| Attribute | Detail |
|-----------|--------|
| Audience | Enterprise/sysadmin |
| Format | NuGet-style .nuspec + install script |
| Notes | More overhead than Scoop/winget for similar audience. Skip unless asked. |

## Blocking Prerequisite: Release Binaries CI

**Without binary artifacts on GitHub Releases, NONE of the above are possible.**

Required GitHub Actions workflow (`release.yml`):
```yaml
on:
  push:
    tags: ['v*']

jobs:
  build:
    strategy:
      matrix:
        include:
          - os: windows-latest
            rid: win-x64
            artifact: QuickSheet-win-x64.zip
          - os: ubuntu-latest
            rid: linux-x64
            artifact: QuickSheet-linux-x64.tar.gz
    runs-on: ${{ matrix.os }}
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with: { dotnet-version: '9.0.x' }
      - run: dotnet publish ExcelConsole.csproj -c Release -r ${{ matrix.rid }} --self-contained -p:PublishSingleFile=true -p:PublishTrimmed=true -o publish/
      - # zip/tar the publish/ directory
      - uses: softprops/action-gh-release@v2
        with:
          files: ${{ matrix.artifact }}
```

**Key publish flags:**
- `--self-contained` — no .NET runtime needed on target machine
- `-p:PublishSingleFile=true` — one executable
- `-p:PublishTrimmed=true` — smaller binary (~30-50MB → ~15-25MB)
- `-r win-x64` / `-r linux-x64` — platform-specific

## Implications for QuickSheet

1. **Immediate next action (Bucket E/B):** Create `.github/workflows/release.yml` that publishes self-contained binaries on tag push. This unblocks ALL package manager submissions. ~40 lines of YAML, touches 1 file, well within the "small change" boundary.

2. **After CI lands:** Submit Scoop manifest to `ScoopInstaller/Extras`. Draft the JSON manifest + PR body now, ship when binaries exist. Scoop's audience (Windows CLI devs) is ideal for QuickSheet.

3. **After Scoop:** Submit winget manifest to `microsoft/winget-pkgs`. Broader reach, same binary.

4. **After winget:** Create `quicksheet-bin` AUR package. Captures the Arch/unixporn audience that the existing `docs/for-students.md` and Reddit drafts target.

5. **README badge:** Once on any package manager, add install badges to README hero section. `scoop install quicksheet` one-liner is far more compelling than "clone and build."

## Effort Estimate

| Step | Effort | Blocked on |
|------|--------|-----------|
| Release CI workflow | 1 run | nothing |
| Scoop manifest draft | 1 run | CI merged + one tagged release |
| winget manifest draft | 1 run | CI merged + one tagged release |
| AUR PKGBUILD draft | 1 run | CI merged + one tagged release |
| README install section | 1 run | at least one package manager accepted |

Total: ~5 runs from "no distribution" to "on 3 package managers with install badges."
