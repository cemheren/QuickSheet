# Agent 10 — fix delivery review findings (Sonnet)

Exact prompt (`$SCRATCH` = session scratch dir).

````text
You are fixing review findings in /home/user/QuickSheet (branch ccr-7ea14cd6-efxsg4; read CLAUDE.md). The changes under review are UNCOMMITTED in the working tree (`git status`, `git diff`; untracked: .github/workflows/libghostty.yml, libghostty.props, scripts/libghostty-smoke/). Do NOT commit or push — pushing publishes an immutable GitHub Release, so everything must be right before that. Another agent is editing SECURITY.md, CONTRIBUTING.md, website/index.html and .claude/skills/grow-quicksheet/SKILL.md in parallel — do NOT touch those. You own: .github/workflows/libghostty.yml, ExcelConsole.csproj, libghostty.props, scripts/build-libghostty.sh, scripts/build-libghostty.ps1, scripts/libghostty-smoke/**, Terminal/GhosttyNative.cs (comment + one small property only).

Context: the workflow builds libghostty-vt 1.0.0 (Ghostty's VT library) from the official pinned source on 4 runners, smoke-tests it via scripts/libghostty-smoke (real Terminal/*.cs binding), and the `publish` job creates the immutable release `libghostty-vt-1.0.0-1`; the csproj target `QsFetchLibGhostty` downloads + SHA-256-verifies the asset for the build. Reference (read-only): libghostty 1.0.0 source at $SCRATCH/ghostty-build/ (do NOT run its build). .NET SDK: $SCRATCH/tools/dotnet/dotnet. Scratch work only under $SCRATCH/fix2/ (never inside the repo except gitignored bin/obj; don't leave stray files like pid files in the repo).

## Fixes

### 1 (MAJOR) Windows must build the GNU ABI
In 1.0.0, src/build/Config.zig:137-141 forces `abi = .msvc` on Windows when no ABI is given → DLL linked against the runner's MSVC runtime (likely imports VCRUNTIME140.dll, absent on clean Windows). Upstream CI uses `-Dtarget=x86_64-windows-gnu`.
- Matrix: win-x64 `zig-target: x86_64-windows-gnu`, win-arm64 `zig-target: aarch64-windows-gnu`. Confirm scripts/build-libghostty.ps1 passes `-Dtarget=$env:ZIG_TARGET` when set (it should already).
- Add an import allowlist check so a bad binary can never be published:
  - New file `scripts/libghostty-smoke/check_deps.py` (Python 3 stdlib only, run with `python3 -I`): mode `pe <dll>` parses the PE import directory (DOS header → PE header → optional header (PE32+) → data directory[1] import table → walk IMAGE_IMPORT_DESCRIPTORs, map RVAs to file offsets via section table, read DLL names) and also delay-load imports (data dir 13) if present; mode `elf <so>` runs nothing external — parse ELF64 dynamic section DT_NEEDED entries (PT_DYNAMIC + DT_STRTAB via section/segment mapping) OR simply shell out to `readelf -d` (available on ubuntu runners) — your call, but stdlib-only parsing is preferred for PE. Allowlists (case-insensitive): PE: kernel32.dll, ntdll.dll, advapi32.dll, ucrtbase.dll, msvcrt.dll, bcrypt.dll, ws2_32.dll, user32.dll, api-ms-win-*.dll; ELF: libc.so.6, libm.so.6, libpthread.so.0, libdl.so.2, librt.so.1, ld-linux-x86-64.so.2, ld-linux-aarch64.so.1. Print all imports; exit 1 listing any not allowed.
  - Test the PE parser locally against a real PE file you can find in the container (e.g. any .dll inside the .NET SDK dir: find /tmp/.../tools/dotnet -name '*.dll' — managed DLLs still have an import table importing mscoree.dll; also find a native one if any, e.g. under tools/dotnet/packs/*win*? If none, managed is fine to prove parsing) and the ELF path against e.g. /usr/lib/x86_64-linux-gnu/libz.so.1 or the dotnet host's libhostfxr.so.
  - Workflow: run it in each build leg after building (Linux: `python3 -I scripts/libghostty-smoke/check_deps.py elf "$lib"`; Windows: `python scripts/libghostty-smoke/check_deps.py pe <dll>` — Windows runners have Python; use `python` there) AND in the publish job over all 4 downloaded artifacts before creating the release (needs checkout — see 5).

### 2 Release notes: glibc 2.28 = Debian 10+, Ubuntu 20.04+, RHEL 8+ (not 18.04).

### 3 Title/consistency
- Title: `"libghostty-vt 1.0.0 (QuickSheet build ${RELEASE_TAG##*-})"` — derive the revision from RELEASE_TAG; put the upstream version (1.0.0) and the source sha256 in workflow-level env vars (`LIBGHOSTTY_VERSION`, `LIBGHOSTTY_SRC_SHA256`) used by the notes.
- In publish (after checkout): fail if `LibGhosttyReleaseTag` in libghostty.props ≠ `$RELEASE_TAG` (extract with grep/sed or python). When the release ALREADY exists: download its SHA256SUMS and, for each of the 4 pins in libghostty.props that is non-empty, fail if it doesn't match (case-insensitive). If the existing release is a draft (`gh release view --json isDraft -q .isDraft` == true) fail with "draft release left over from a failed publish; delete it and re-run".

### 4 Triggers/races
- `push:` add `branches: ['**']` (prevents tag pushes from triggering).
- `publish`: `if: github.event_name != 'pull_request' && github.ref_type == 'branch'`; add `concurrency: { group: libghostty-publish, cancel-in-progress: false }`.
- Add `timeout-minutes: 60` to build legs and 15 to publish.

### 5 Pin actions by full commit SHA + persist-credentials false
- Resolve the latest v4 release commit SHA of actions/checkout, actions/setup-dotnet, actions/upload-artifact, actions/download-artifact with `git ls-remote --tags https://github.com/actions/<name>` (pick the highest v4.x.y tag; use the peeled `^{}` commit if present). Write `uses: actions/checkout@<sha> # v4.x.y`.
- `actions/checkout` everywhere with `persist-credentials: false` (build legs run third-party build.zig code). Publish job checks out too (for libghostty.props and check_deps.py) — still persist-credentials false; `gh` uses GH_TOKEN env.

### 6 License/notice assets
The release must carry license notices for what's linked into the binaries (libghostty MIT + third-party deps listed in the source tarball's `deps/libghostty-vt-third-party.json` if present, plus the top-level LICENSE).
- build-libghostty.sh/.ps1: after building, copy `$SRC_DIR/LICENSE` → `$DEST_DIR/libghostty-vt-LICENSE.txt` and, if it exists, `$SRC_DIR/deps/libghostty-vt-third-party.json` → `$DEST_DIR/libghostty-vt-third-party.json` (print a note if absent). (Check ghostty-build/ for the actual file names: LICENSE at repo root; deps/libghostty-vt-third-party.json.)
- Workflow: the linux-x64 leg additionally uploads those two files in a separate artifact `notices` (if-no-files-found: warn for the json, but LICENSE must exist → error); publish attaches them as release assets `libghostty-vt-LICENSE.txt` and `libghostty-vt-third-party.json` (the json only if present) and includes them in SHA256SUMS. These are NOT downloaded by the csproj.

### 7 Download timeout
`DownloadFile ... Timeout="20000"` (ms).

### 9 -warnaserror must not break offline builds
- Add `<MSBuildWarningsAsMessages>$(MSBuildWarningsAsMessages);MSB3923</MSBuildWarningsAsMessages>` (DownloadFile failure) and turn the "Could not download libghostty-vt" Warning into `<Message Importance="high" .../>`. KEEP the SHA-256 mismatch as a Warning (security signal; failing a -warnaserror build on tampering is correct).
- Verify: offline build (pin set to some value, base URL pointing at a closed localhost port) with `-warnaserror` succeeds.

### 10 RID OS + host arch
- When `$(RuntimeIdentifier)` is set: OS from its prefix — `win-*` → windows, `linux-musl-*`/`linux-bionic-*` → unsupported (no native lib, no warning), `linux-*` → linux, anything else → unsupported. Arch = part after last '-'.
- When not set: OS from the host (as now) and arch from `$(NETCoreSdkRuntimeIdentifier)` (part after last '-'), falling back to OSArchitecture if empty.
- Verify with -getProperty: `-r win-x64` on Linux → QsNativeRid=win-x64 / QsNativeFileName=ghostty-vt.dll / asset ghostty-vt-win-x64.dll; `-r linux-musl-arm64` → empty QsNativeRid; no -r → linux-x64 here.
- Make sure the local-override ItemGroup and target conditions still work with the new properties.

### 11 Smoke test stricter
- Add to Terminal/GhosttyNative.cs a `public static string? LoadedPath` set by ProbeLoad to the path/name that actually loaded (no behavior change otherwise). Also update the class header comment: library is downloaded by `dotnet build` from this repo's CI-built release (pinned SHA-256 in libghostty.props), or a local `native/<rid>/` build, or QUICKSHEET_LIBGHOSTTY.
- Smoke: assert `GhosttyNative.Version == "1.0.0"` and that `LoadedPath` equals the QUICKSHEET_LIBGHOSTTY value (Path.GetFullPath compare, case-insensitive on Windows).

## Verify
- `python3 -I -c "import yaml..."` parse of the workflow (PyYAML available? if not, `pip install --user pyyaml` into scratch or skip and say so); `bash -n` on the .sh and on extracted bash run blocks.
- `$DOTNET build ExcelConsole.csproj` → 0 warnings/0 errors (empty pins message). Smoke project builds; clean its bin/obj.
- csproj scenarios with a local `python3 -m http.server --bind 127.0.0.1` from the scratch dir: good pin downloads; offline + -warnaserror succeeds; -getProperty checks from 10. Kill the server; remove native/ and any test leftovers; `git status --short` should show only expected files.

Report: per finding what changed (file:line), the resolved action SHAs + tags, verification output, deviations.
````
