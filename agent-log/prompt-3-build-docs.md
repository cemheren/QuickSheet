# Agent 3 — build scripts, csproj, docs (Haiku)

Exact prompt given by the orchestrator (`$SCRATCH` = session scratch dir, outside the repo).

````text
You are editing the repo at /home/user/QuickSheet (a .NET 9 desktop-wallpaper spreadsheet app; read /home/user/QuickSheet/CLAUDE.md first). Do NOT commit. Other agents are concurrently editing ConPtyProcess.cs, Platform/Linux/PtyProcess.cs, InlineProcessManager.cs, and creating Terminal/*.cs — do NOT touch those. You own ONLY: scripts/build-libghostty.sh (new), scripts/build-libghostty.ps1 (new), ExcelConsole.csproj, .gitignore, README.md, CONTRIBUTING.md, CLAUDE.md, CHANGELOG.md, docs/faq.md, docs/tour.md.

## Context
We are migrating inline-process console emulation (ANSI-regex stripping of PTY output) to libghostty-vt — the VT emulator C library from the Ghostty terminal (stable C ABI 1.0, release 1.0.0). It's a NATIVE library (not NuGet — the project's zero-NuGet rule still holds). It's loaded at runtime via hand-written P/Invoke from `ExcelConsole.Terminal.GhosttyNative` which looks for, in order: env var `QUICKSHEET_LIBGHOSTTY` (full path), then next to the executable (`ghostty-vt.dll` on Windows; `libghostty-vt.so.1` / `libghostty-vt.so` on Linux), then system library paths. If it's not found, QuickSheet falls back to the old regex stripper (`LegacyVtScreen`) and prints a one-line note to stderr. So building without Zig still works.

## 1. scripts/build-libghostty.sh (bash, `set -euo pipefail`, chmod +x)
Builds libghostty-vt 1.0.0 from the OFFICIAL signed source tarballs, verifying pinned SHA-256 checksums, fully offline build (deps tarball passed via --system), and copies the shared library to `<repo>/native/linux-x64/` (derive the rid: x86_64→linux-x64, aarch64→linux-arm64).
Facts (verified):
- Source: https://release.files.ghostty.org/libghostty-vt/1.0.0/libghostty-vt-1.0.0.tar.gz  sha256 4a6ef758be721923e1969e7b81b443dcc92d8dd3fb5a39600c03a8817a753676 → unpacks to dir `libghostty-vt-1.0.0`
- Deps: https://release.files.ghostty.org/libghostty-vt/1.0.0/libghostty-vt-1.0.0-deps.tar.gz  sha256 4d5ea5e448e4c7ea8fb55ccbaa8668f2a97d44f438cee831e2b7ecd162d85c29 → unpacks to dir `libghostty-vt-1.0.0-deps`
- Each file also has a `.minisig` (minisign public key `RWQlAjJC23149WL2sEpT/l0QKy7hMIFhYdQOFy0Z7z7PbneUgvlsnYcV`). If `minisign` is on PATH, ALSO verify signatures (`minisign -V -P <key> -m <file>`; download `<url>.minisig`); if not, print a note that the pinned sha256 was checked but the signature was skipped.
- Requires Zig 0.17.x exactly (check `zig version` starts with `0.17.`; error out with a helpful message + https://ziglang.org/download/ otherwise).
- Build command (run inside the unpacked source dir, it's a source tarball so -Demit-lib-vt is not needed but harmless — include it):
  `zig build --prefix "$STAGE" --system "$WORK/libghostty-vt-1.0.0-deps" -Demit-lib-vt -Doptimize=ReleaseFast -Dcpu=baseline -Dstrip=true`
  with `ZIG_GLOBAL_CACHE_DIR="$WORK/zig-cache"`.
- On Linux the outputs in `$STAGE/lib/` are `libghostty-vt.so.1.0.0` (real file), `libghostty-vt.so.1` and `libghostty-vt.so` (symlinks). Copy the real file to `native/<rid>/libghostty-vt.so.1` (dereference; `cp -L`) — a single file, no symlinks in our tree.
- Use a temp work dir (mktemp -d) cleaned up via trap; download with curl -fsSL; verify with `sha256sum -c` (or `shasum -a 256` fallback for macOS).
- Variables at top: LIBGHOSTTY_VERSION, SRC_SHA256, DEPS_SHA256, MINISIGN_PUBKEY, so bumping is a one-place change. Print where the library was installed and remind the user to rebuild (`dotnet build`) so it's copied to the output dir.

## 2. scripts/build-libghostty.ps1 (PowerShell 5.1+/7 compatible)
Same thing for Windows: download with Invoke-WebRequest, verify with Get-FileHash -Algorithm SHA256 against the pinned values, optional minisign if found, `tar -xzf` (tar.exe ships with Windows 10+), check `zig version` 0.17.x, same zig build flags, output DLL is `$STAGE\bin\ghostty-vt.dll` → copy to `<repo>\native\win-x64\ghostty-vt.dll` (rid from $env:PROCESSOR_ARCHITECTURE: AMD64→win-x64, ARM64→win-arm64). `$ErrorActionPreference = 'Stop'`, temp dir cleaned up in finally.

## 3. ExcelConsole.csproj
Add an ItemGroup that copies any prebuilt native lib from `native/**` to the output directory root (flattened), only if present — e.g.:
```xml
<!-- libghostty-vt native library (optional; built by scripts/build-libghostty.*). Not a NuGet package. -->
<ItemGroup>
  <None Include="native/**/ghostty-vt.dll;native/**/libghostty-vt.so*;native/**/libghostty-vt*.dylib"
        CopyToOutputDirectory="PreserveNewest" Link="%(Filename)%(Extension)" Visible="false" />
</ItemGroup>
```
Make it OS-appropriate: on Windows only copy from `native/win-*`, on Linux only from `native/linux-*` matching the current architecture if easy (use `$([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture)` → X64/Arm64 to pick `x64`/`arm64`), so a dev with both folders doesn't get the wrong one. Wildcards that match nothing must not break the build. Also make sure the existing `Compile Remove` items still work and that files under native/ and scripts/ are never compiled.
Verify with the .NET SDK at $SCRATCH/tools/dotnet/dotnet: create a dummy file `native/linux-x64/libghostty-vt.so.1` (any content), run `$DOTNET build /home/user/QuickSheet/ExcelConsole.csproj`, confirm it's copied to `bin/Debug/net9.0/libghostty-vt.so.1`, then DELETE the dummy file and the native/ dir. (Compile errors in ConPtyProcess.cs/PtyProcess.cs/InlineProcessManager.cs/Terminal/* are from other agents mid-edit — ignore those, only check the copy works; if the build fails before copying, use `$DOTNET msbuild ExcelConsole.csproj -t:_CopyOutOfDateSourceItemsToOutputDirectory` or just inspect `-getItem:None` with `$DOTNET msbuild -getItem:None` to confirm the item is picked up.)

## 4. .gitignore
Add `native/` (built binaries are never committed — supply-chain policy: build from pinned, checksummed source).

## 5. Docs (keep edits concise and in the existing voice)
- README.md line ~60 and CONTRIBUTING.md line ~23, docs/faq.md line ~15, docs/tour.md line ~103 mention "zero NuGet dependencies — all native interop (X11, WinForms, ConPTY) is hand-written P/Invoke". Keep that claim (still true) and add that inline-command output is rendered by **libghostty-vt** (Ghostty's terminal emulator library), an optional native library built from a pinned, checksum-verified source tarball with `scripts/build-libghostty.sh` / `.ps1` (needs Zig 0.17); without it QuickSheet falls back to basic ANSI stripping. Don't overdo it — one or two sentences per place; in README add a short "Optional: libghostty-vt" subsection near the build/run instructions (find them) with the two script commands.
- CONTRIBUTING.md: in the rules table or nearby, note that native libraries must be built from pinned+checksummed sources by a script, never committed as binaries, and must stay optional.
- CLAUDE.md: (a) Build & Run: add `./scripts/build-libghostty.sh   # optional: build libghostty-vt (Zig 0.17) into native/`; (b) Architecture: add a bullet `- **`Terminal/`** — `IVtScreen` screen model for PTY output. `GhosttyVtScreen` wraps libghostty-vt (hand-written P/Invoke in `GhosttyNative`, loaded from `QUICKSHEET_LIBGHOSTTY`, next to the exe, or system path); `LegacyVtScreen` is the regex fallback when the lib is absent. `VtScreen.Create` picks.` and update the `InlineProcessManager` bullet to say PTY bytes are fed into an `IVtScreen`; (c) Conventions: add namespace `ExcelConsole.Terminal`, and a bullet that native libs (libghostty-vt) are optional, built from pinned checksummed source, never committed.
- CHANGELOG.md: read the format; add an "Unreleased" section at the top (if there isn't one) with a "Changed" entry: inline `i:` command output now runs through libghostty-vt, so cursor movement, clear-screen, progress bars (`\r`), and full-screen TUIs show the real final screen + scrollback instead of regex-stripped text; also fixes UTF-8 characters split across reads. Optional native dep; falls back to the old stripper when missing.

Report back: list of every file changed, the exact csproj snippet you added, the build-copy verification output, and `bash -n scripts/build-libghostty.sh` result.
````
