# Agent 6 — CI workflow building libghostty-vt (Sonnet)

Exact prompt (`$SCRATCH` = session scratch dir).

````text
You are working in /home/user/QuickSheet (C#/.NET 9 desktop-wallpaper spreadsheet; read /home/user/QuickSheet/CLAUDE.md first). Branch ccr-7ea14cd6-efxsg4. Do NOT commit or push. Other agents are editing ExcelConsole.csproj, libghostty.props and the docs in parallel — you own ONLY: `.github/workflows/libghostty.yml` (new), `scripts/build-libghostty.sh`, `scripts/build-libghostty.ps1`, and a new folder `scripts/libghostty-smoke/`.

## Background
QuickSheet now renders inline-command output with libghostty-vt (Ghostty's VT library, release 1.0.0), via hand-written P/Invoke in Terminal/GhosttyNative.cs + Terminal/GhosttyVtScreen.cs, with a regex fallback (Terminal/LegacyVtScreen.cs). Ghostty publishes NO prebuilt Windows/Linux libraries, only source tarballs. Users install from source (`git clone` → `dotnet run`) and must not need Zig. Decision: QuickSheet's own GitHub Actions builds libghostty-vt from the official, SHA-256-pinned source tarball (exactly what scripts/build-libghostty.{sh,ps1} already do — read both), publishes the binaries to an immutable GitHub Release of this repo, and `dotnet build` downloads them with pinned hashes (another agent does that csproj part).

Shared contract (other agents rely on these EXACT names):
- Release tag: `libghostty-vt-1.0.0-1` (the `-1` is a QuickSheet build revision; bump it to rebuild). Put it in a workflow-level `env: RELEASE_TAG`.
- Release assets (flat names):
  - `libghostty-vt-linux-x64.so.1`
  - `libghostty-vt-linux-arm64.so.1`
  - `ghostty-vt-win-x64.dll`
  - `ghostty-vt-win-arm64.dll`
  - `SHA256SUMS` (sha256sum format, `<hex>  <asset>` for the four files)
- Download URL pattern: `https://github.com/cemheren/QuickSheet/releases/download/<tag>/<asset>`

## 1. Build scripts: add an optional Zig target override
- scripts/build-libghostty.sh: if env `ZIG_TARGET` is set and non-empty, pass `-Dtarget="$ZIG_TARGET"` to `zig build`. Also allow env `LIBGHOSTTY_OUT` to override the output directory (default stays `$REPO_ROOT/native/$RID`). Keep everything else. Keep `bash -n` clean.
- scripts/build-libghostty.ps1: same two env vars ($env:ZIG_TARGET, $env:LIBGHOSTTY_OUT). Read the script carefully — it has never been run; fix any obvious bugs you see (PowerShell 5.1 + 7 compatible; `tar.exe` usage; Get-FileHash comparisons are case-insensitive hex; `$LASTEXITCODE` must be checked after native commands like zig/tar/minisign because $ErrorActionPreference='Stop' does NOT catch native non-zero exits).

## 2. Smoke test project: scripts/libghostty-smoke/
A tiny console app that loads a given libghostty-vt build through QuickSheet's REAL binding code and checks it works, so CI proves the P/Invoke is right on every platform before publishing.
- `scripts/libghostty-smoke/libghostty-smoke.csproj`: OutputType Exe, net9.0, Nullable+ImplicitUsings enable, `EnableDefaultCompileItems=false`, `<Compile Include="../../Terminal/*.cs" />` plus `Program.cs`. If Terminal/GhosttyVtScreen.cs uses `#if PLATFORM_WINDOWS`, define PLATFORM_WINDOWS when `$([MSBuild]::IsOSPlatform('Windows'))` and PLATFORM_LINUX on Linux (mirror ExcelConsole.csproj).
- `Program.cs` (namespace-free top-level or `ExcelConsole.Terminal` usage — types are `internal`, which is fine since they're compiled into this assembly):
  - Require env `QUICKSHEET_LIBGHOSTTY` to be set (absolute path) — exit 2 with message otherwise.
  - Assert `GhosttyNative.IsAvailable`; print `GhosttyNative.Version`; on failure print `GhosttyNative.LoadError` and exit 1.
  - `using var s = VtScreen.Create(40, 5, 100);` assert `s is GhosttyVtScreen` (not legacy).
  - Feed bytes in SEVERAL small chunks (split a UTF-8 char and an escape sequence across Write calls): `"\x1b[31mred\x1b[0m é\r\n"`, then `"progress 10%\rprogress 99%\r\n"`, then `"to be erased\x1b[2K\rkept\r\n"`. Assert GetText() == "red é\nprogress 99%\nkept".
  - Assert IsDirty false right after GetText, true after another Write.
  - Clear screen + home: write `"\x1b[2J\x1b[Hfresh"`; assert GetText() ends with "fresh".
  - Scrollback: on a new 20x3 screen write 50 lines `line N\r\n`; assert GetText contains "line 0" and "line 49" (scrollback kept, default formatter includes history).
  - Resize to 10x3 and back without exceptions; Dispose; call Write after Dispose (must be a no-op, no crash).
  - Print PASS/FAIL per check; exit 0 only if all pass.
- Read Terminal/*.cs to use the real APIs (IVtScreen: Write(ReadOnlySpan<byte>), Resize, IsDirty, GetText; VtScreen.Create; GhosttyNative.IsAvailable/Version/LoadError).
- ExcelConsole.csproj already has `<Compile Remove="native/**;scripts/**" />`, so this project won't be compiled into the app. Check `quicksheet.sln` doesn't need changes (don't add the smoke project to it).
- Verify locally: `$DOTNET build scripts/libghostty-smoke/libghostty-smoke.csproj` builds clean, and running it WITHOUT the env var exits 2, and with `QUICKSHEET_LIBGHOSTTY=/nonexistent/libghostty-vt.so.1` exits 1 with a clear message. .NET SDK: $SCRATCH/tools/dotnet/dotnet. Clean up bin/obj under scripts/libghostty-smoke after (they're gitignored anyway, but don't leave stray files). Also make sure `$DOTNET build ExcelConsole.csproj` at repo root still builds (no duplicate Program.cs pickup).

## 3. .github/workflows/libghostty.yml
Name: `libghostty-vt`. Supply-chain stance: use only first-party GitHub actions (`actions/checkout@v4`, `actions/setup-dotnet@v4`, `actions/upload-artifact@v4`, `actions/download-artifact@v4`) — NO third-party actions. Pin them by major tag as written (fine for GitHub-owned actions).
Triggers:
- `push` with `paths: [ 'scripts/build-libghostty.*', 'scripts/libghostty-smoke/**', 'Terminal/**', '.github/workflows/libghostty.yml' ]` (any branch).
- `pull_request` with the same paths (build + smoke only; never publish).
- `workflow_dispatch`.
Top-level `permissions: contents: read`.

Job `build` (matrix, `fail-fast: false`):
| rid | runs-on | zig asset | zig sha256 | ZIG_TARGET | script |
|---|---|---|---|---|---|
| linux-x64 | ubuntu-24.04 | zig-x86_64-linux-0.17.0.tar.xz | 1cbe9df9f27e6b78d14ccbca43b6703a404ef79ef1c463de901d7f088d4e2026 | x86_64-linux-gnu.2.28 | build-libghostty.sh |
| linux-arm64 | ubuntu-24.04-arm | zig-aarch64-linux-0.17.0.tar.xz | 9e8d11661d4ae3bd57702a3832781e23ad151dde5798e16a5ccd503f65234ff8 | aarch64-linux-gnu.2.28 | build-libghostty.sh |
| win-x64 | windows-2025 | zig-x86_64-windows-0.17.0.zip | b5663f69581dcf391293fbf16c06cb80d81d806545ce618b4d0bab7f0eb8c428 | (empty) | build-libghostty.ps1 |
| win-arm64 | windows-11-arm | zig-aarch64-windows-0.17.0.zip | 0a59d91fa1cb40cf068e9b0954434ce973500c7a2ea749f1e01af62cdab52d26 | (empty) | build-libghostty.ps1 |
Zig URL: `https://ziglang.org/download/0.17.0/<asset>`. Download, verify sha256 (fail on mismatch), extract, add the extracted dir to `$GITHUB_PATH`. (glibc 2.28 pin = runs on Debian 10+/Ubuntu 18.04+/RHEL 8+.)
Steps per matrix entry:
1. checkout
2. install Zig (bash on Linux; `shell: pwsh` on Windows with Invoke-WebRequest + Get-FileHash + Expand-Archive)
3. run the build script with `LIBGHOSTTY_OUT=${{ runner.temp }}/out` and `ZIG_TARGET` from the matrix (Windows: `pwsh -File scripts/build-libghostty.ps1`; also try under Windows PowerShell 5.1? not needed — pwsh is fine)
4. Linux only: sanity checks — `readelf --dyn-syms` shows `ghostty_terminal_new`; the highest `GLIBC_x.y` version required (`objdump -T` or `readelf -V`) is ≤ 2.28, fail otherwise.
5. `actions/setup-dotnet@v4` with `dotnet-version: 9.0.x`; run the smoke test: `dotnet run --project scripts/libghostty-smoke/libghostty-smoke.csproj -c Release` with env `QUICKSHEET_LIBGHOSTTY` = the absolute path of the built library.
6. Rename the built file to its asset name (contract above) and `upload-artifact` it (name = rid, `if-no-files-found: error`, retention 14 days).

Job `publish`: `needs: build`; runs on ubuntu-24.04; `if: github.event_name != 'pull_request'`; `permissions: contents: write`; env `GH_TOKEN: ${{ github.token }}`.
- Checkout NOT needed; use `gh` with `--repo ${{ github.repository }}`.
- If `gh release view "$RELEASE_TAG"` succeeds → print "release exists; immutable, not overwriting" and exit 0 (we never replace published binaries; pinned hashes in the repo would break). Otherwise download all artifacts, `sha256sum` them into SHA256SUMS, and `gh release create "$RELEASE_TAG" --target "$GITHUB_SHA" --title "libghostty-vt 1.0.0 (QuickSheet build 1)" --notes "<notes>" <4 files> SHA256SUMS`. Notes should say: built by this workflow from the official libghostty-vt 1.0.0 source tarball (sha256 4a6ef758…), offline with the deps tarball, Zig 0.17.0, `-Doptimize=ReleaseFast -Dcpu=baseline -Dstrip=true`, Linux glibc ≥ 2.28; link to the workflow run (`$GITHUB_SERVER_URL/$GITHUB_REPOSITORY/actions/runs/$GITHUB_RUN_ID`). Mark it `--prerelease=false`, and NOT `--latest` (use `--latest=false`) so the app's own releases stay "latest".
- In BOTH cases (new or existing release), write to `$GITHUB_STEP_SUMMARY` a ready-to-paste MSBuild snippet with the hashes (read from the release's SHA256SUMS via `gh release download "$RELEASE_TAG" -p SHA256SUMS` when it already existed):
```xml
<LibGhosttySha256_linux_x64>…</LibGhosttySha256_linux_x64>
<LibGhosttySha256_linux_arm64>…</LibGhosttySha256_linux_arm64>
<LibGhosttySha256_win_x64>…</LibGhosttySha256_win_x64>
<LibGhosttySha256_win_arm64>…</LibGhosttySha256_win_arm64>
```
Avoid `${{ }}` interpolation of anything attacker-controllable inside `run:` scripts (use env vars); there's nothing user-controlled here except refs, but keep the habit.

Validate the YAML parses: `python3 -I -c "import yaml,sys; yaml.safe_load(open(sys.argv[1]))" .github/workflows/libghostty.yml` (if PyYAML is missing, `pip install --user pyyaml` is fine, or skip and say so). `bash -n` both the .sh script and the bash `run:` blocks you can extract.

Report back: files changed, the final workflow YAML in full, the smoke test checks, verification output, and any concerns (e.g. runner availability of ubuntu-24.04-arm / windows-11-arm for public repos, whether zig's Windows default ABI produces a DLL without extra runtime deps).
````
