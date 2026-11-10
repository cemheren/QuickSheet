# Agent 6 — CI workflow building libghostty-vt (Sonnet) — report

Everything you asked for is written, but the real library has not been run. The smoke test and the ps1 changes are untested against a real build, and `pwsh` isn't installed here, so `build-libghostty.ps1` was only reviewed, not executed.

**Verified locally**
- `bash -n` passes on `build-libghostty.sh` and on all 6 extracted bash `run:` blocks.
- The workflow YAML parses with PyYAML.
- `dotnet build scripts/libghostty-smoke/libghostty-smoke.csproj` builds clean (0 warnings, 0 errors).
- Running the smoke test without `QUICKSHEET_LIBGHOSTTY` exits 2 with a clear message.
- With `QUICKSHEET_LIBGHOSTTY=/nonexistent/libghostty-vt.so.1` it exits 1: `FAIL: libghostty-vt could not be loaded: native library not found (...)`.
- `dotnet build ExcelConsole.csproj` at the repo root still builds (0 warnings, 0 errors), so no duplicate Program.cs is picked up.
- I removed the smoke project's `bin`/`obj`. `quicksheet.sln` needs no change; I did not add the smoke project to it.

**Files changed** (`.github/workflows/` and `scripts/libghostty-smoke/` are new)
- `/home/user/QuickSheet/.github/workflows/libghostty.yml`
- `/home/user/QuickSheet/scripts/build-libghostty.sh`
- `/home/user/QuickSheet/scripts/build-libghostty.ps1`
- `/home/user/QuickSheet/scripts/libghostty-smoke/libghostty-smoke.csproj`
- `/home/user/QuickSheet/scripts/libghostty-smoke/Program.cs`

**Build scripts**
- `build-libghostty.sh`: `ZIG_TARGET`, when non-empty, adds `-Dtarget=...` to `zig build` (the args now live in a bash array). `LIBGHOSTTY_OUT` overrides the output dir via `DEST_DIR="${LIBGHOSTTY_OUT:-$REPO_ROOT/native/$RID}"`.
- `build-libghostty.ps1`: the same two env vars. I also fixed these bugs:
  - `Fail` used `Write-Error`, which under `$ErrorActionPreference='Stop'` throws before `exit 1` runs. It now writes to stderr and exits 1.
  - The hash comparison is now `-ine`, so it is case-insensitive.
  - `$ProgressPreference='SilentlyContinue'`, because `Invoke-WebRequest` is very slow on 5.1 with progress on.
  - It uses `System32\tar.exe` (bsdtar) first, because Git's GNU tar can misread `C:\` paths.
  - The architecture is read from `PROCESSOR_ARCHITEW6432` first.
  - If `bin\ghostty-vt.dll` isn't where expected, it searches the stage dir for `ghostty-vt*.dll`.
  - `$LASTEXITCODE` is checked after `zig` and `tar` as before; I did not add a check after the Zig call beyond what was already there.

**Smoke test checks** (the csproj defines `PLATFORM_WINDOWS` or `PLATFORM_LINUX` like `ExcelConsole.csproj`; `Program.cs` compiles `Terminal/*.cs` directly)
1. `GhosttyNative.IsAvailable` and `Version`. On failure it prints `LoadError` and exits 1.
2. `VtScreen.Create(40,5,100)` returns a `GhosttyVtScreen`, not the legacy fallback.
3. Chunked writes: the first chunk is cut inside `ESC[31m` and again between the two bytes of `é`, the second and third inputs are also split. `GetText()` must equal `"red é\nprogress 99%\nkept"`.
4. `IsDirty` is false after `GetText` and true after another `Write`.
5. `ESC[2J ESC[H fresh`: `GetText()` ends with `fresh`.
6. `Resize(10,3)` then `Resize(40,5)` throw nothing.
7. `Write` and `Resize` after `Dispose`, plus a double dispose from `using`, must not crash.
8. On a 20x3 screen with 50 lines written, `GetText()` contains `line 0\n` and `line 49`.

It prints PASS/FAIL per check and exits 0 only if everything passes. Any unexpected exception is reported as a FAIL.

**Workflow**
- The YAML is in the file above. I could not paste the full text, so it is not repeated here.
- Triggers are `push` and `pull_request` on the four paths (duplicated, not YAML anchors, to be safe on GitHub) plus `workflow_dispatch`.
- `contents: read` at top level and `RELEASE_TAG: libghostty-vt-1.0.0-1` in a workflow-level `env`.
- `build` is a 4-entry matrix with `fail-fast: false`, using your Zig assets, hashes and targets. Matrix values go through `env`, not inline `${{ }}` in scripts.
- Steps per entry:
  - Install Zig and verify its sha256.
  - Build with `LIBGHOSTTY_OUT=$RUNNER_TEMP/out`.
  - On Linux only, check that `ghostty_terminal_new` is exported and the highest `GLIBC_` version is at most 2.28.
  - Run the smoke test under `setup-dotnet` 9.0.x.
  - Copy the library to its contract asset name and `upload-artifact` it (14-day retention, `if-no-files-found: error`).
- `publish`:
  - It needs `build`, is skipped on `pull_request`, and has `contents: write`.
  - If the release exists it prints "immutable, not overwriting" and downloads `SHA256SUMS`.
  - Otherwise it builds `SHA256SUMS` for the 4 assets and runs `gh release create --target $GITHUB_SHA --prerelease=false --latest=false`.
  - In both cases it writes the 4 `<LibGhosttySha256_*>` lines to `$GITHUB_STEP_SUMMARY`.
- Only `actions/checkout`, `setup-dotnet`, `upload-artifact` and `download-artifact` are used.

**Concerns**
- **Unverified build output:** the scripts have never been run on Windows, and I couldn't run Zig or the real library here. I don't know that Zig's Windows build puts `ghostty-vt.dll` at `bin\`; the ps1 fallback search covers other paths. The first CI run will prove the build, the P/Invoke and the smoke checks. In particular, check that `GetText()` really returns `"red é\nprogress 99%\nkept"` on the real library, since I wrote that expectation from your spec.
- **Windows DLL dependencies:** Zig's default Windows target ABI is gnu (mingw-style, no MSVC runtime), so I expect no extra runtime DLLs. It is not checked, so I'd add a `dumpbin /dependents` step if you want it enforced.
- **ARM runners:** `ubuntu-24.04-arm` and `windows-11-arm` are free for public repos but have limited capacity. The Windows ARM64 job also needs a native ARM .NET 9 SDK from `setup-dotnet`, which I believe is available. If either runner is unavailable, that matrix entry fails and `publish` is skipped, since `needs: build` requires all of them.
- **Release gating:** `publish` runs on any non-PR push to any branch that touches the watched paths, and it creates the release if the tag doesn't exist. `--target $GITHUB_SHA` points the tag at that commit. If you want releases only from main, add a ref condition to `publish`.
