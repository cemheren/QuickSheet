# Agent 7 — csproj pinned download (Sonnet)

Exact prompt (`$SCRATCH` = session scratch dir).

````text
You are working in /home/user/QuickSheet (C#/.NET 9 desktop-wallpaper spreadsheet; read /home/user/QuickSheet/CLAUDE.md first). Branch ccr-7ea14cd6-efxsg4. Do NOT commit or push. Other agents are editing .github/workflows/, scripts/ and docs in parallel — you own ONLY `ExcelConsole.csproj` and a new file `libghostty.props` at the repo root.

## Background
QuickSheet renders inline-command output with libghostty-vt (Ghostty's VT C library, release 1.0.0) via hand-written P/Invoke (Terminal/GhosttyNative.cs). The loader looks for `QUICKSHEET_LIBGHOSTTY` (absolute path), then `ghostty-vt.dll` (Windows) / `libghostty-vt.so.1` or `libghostty-vt.so` (Linux) in AppContext.BaseDirectory, then (Linux only) system paths; if absent it falls back to a regex stripper. Today the library only appears next to the exe if a contributor ran `scripts/build-libghostty.sh|.ps1` (needs Zig) which puts it in `native/<rid>/`; ExcelConsole.csproj copies it (read the current csproj: QsNativeArch property + two ItemGroups).

New decision: users must NOT need Zig. QuickSheet's own CI (another agent is writing it) builds the library from Ghostty's official pinned source and publishes it as GitHub Release assets. `dotnet build` must download the right binary automatically and verify a SHA-256 pinned in the repo.

Contract:
- Release tag `libghostty-vt-1.0.0-1`; base URL `https://github.com/cemheren/QuickSheet/releases/download/libghostty-vt-1.0.0-1/`
- Assets: `libghostty-vt-linux-x64.so.1`, `libghostty-vt-linux-arm64.so.1`, `ghostty-vt-win-x64.dll`, `ghostty-vt-win-arm64.dll`
- In the output dir they must be named `libghostty-vt.so.1` (Linux) / `ghostty-vt.dll` (Windows).
- Hash pins are NOT known yet (CI hasn't run). Leave them as empty properties; with an empty pin the download is skipped with a single build message (Importance=high, NOT a warning — warnings would show on every build until pinned... actually use a warning only if a pin is set but verification fails).

## libghostty.props (new, imported by the csproj)
```xml
<Project>
  <!-- libghostty-vt: QuickSheet's single external dependency. Built by .github/workflows/libghostty.yml
       from Ghostty's official, SHA-256-pinned source tarball and published as release assets of this repo.
       Bump LibGhosttyReleaseTag + the four hashes together (the workflow's run summary prints them). -->
  <PropertyGroup>
    <LibGhosttyReleaseTag>libghostty-vt-1.0.0-1</LibGhosttyReleaseTag>
    <LibGhosttyBaseUrl>https://github.com/cemheren/QuickSheet/releases/download/$(LibGhosttyReleaseTag)/</LibGhosttyBaseUrl>
    <LibGhosttySha256_linux_x64></LibGhosttySha256_linux_x64>
    <LibGhosttySha256_linux_arm64></LibGhosttySha256_linux_arm64>
    <LibGhosttySha256_win_x64></LibGhosttySha256_win_x64>
    <LibGhosttySha256_win_arm64></LibGhosttySha256_win_arm64>
  </PropertyGroup>
</Project>
```
(Adjust comments as needed; keep property names exactly.)

## ExcelConsole.csproj
Replace the existing native-copy ItemGroups with a target-based flow (keep the existing `Compile Remove`/`None Remove` for native/** and scripts/**, and keep the QsNativeArch logic — RID-derived else host arch, lowercase):
1. `<Import Project="libghostty.props" />`.
2. Compute `QsNativeRid` = `win-$(QsNativeArch)` on Windows, `linux-$(QsNativeArch)` on Linux, empty otherwise. Compute `QsNativeFileName` (`ghostty-vt.dll` / `libghostty-vt.so.1`), `QsNativeAsset` (contract names), and `QsNativeSha256` = the matching `LibGhosttySha256_*` property (rid with '-' → '_').
3. Opt-out property `QuickSheetLibGhosttyDownload` defaulting to `true`; `-p:QuickSheetLibGhosttyDownload=false` disables downloading (offline / paranoid builds). Also skip automatically when `$(DesignTimeBuild)` is true.
4. Local override wins: if `native/$(QsNativeRid)/$(QsNativeFileName)` exists (from the build script), use that and don't download.
5. Otherwise a target `QsFetchLibGhostty` (BeforeTargets="AssignTargetPaths" or whatever reliably precedes the copy-to-output machinery — verify empirically) that:
   - Cache path: `$(MSBuildProjectDirectory)/obj/libghostty/$(LibGhosttyReleaseTag)/$(QsNativeRid)/$(QsNativeFileName)` (use `$(BaseIntermediateOutputPath)` if set). Download only if the cached file is missing.
   - Use the built-in MSBuild `DownloadFile` task (DestinationFolder + DestinationFileName = the asset name, then Copy/Move to the cache name) with `ContinueOnError="true"` and `Retries="2"`; never fail the build because of network problems — emit a warning: "Could not download libghostty-vt (…); inline commands will use basic ANSI stripping. Build offline with -p:QuickSheetLibGhosttyDownload=false to silence."
   - Verify with the built-in `VerifyFileHash` task (Algorithm SHA256, Hash = pin). On mismatch: delete the downloaded file, emit a WARNING about the mismatch (security-relevant, so warning is right) and do not copy it. Do the download into a temp name and only move to the cache path after verification passes, so a bad file is never cached. (VerifyFileHash with ContinueOnError="true" + check `$(MSBuildLastTaskResult)`.)
   - Also re-verify a cached file each build? Cheap enough (a few MB) — yes, verify every build; if the cached file fails, delete it and re-download once.
   - On success add an item `<None Include="<cache path>" Link="$(QsNativeFileName)" CopyToOutputDirectory="PreserveNewest" CopyToPublishDirectory="PreserveNewest" Visible="false" />` so it lands at the root of both build and publish output.
   - If no pin is set for this rid (empty), skip the download with `<Message Importance="high" Text="libghostty-vt: no pinned hash for $(QsNativeRid) yet; skipping download (basic ANSI stripping will be used)." />`.
6. The local-override item should also carry `CopyToPublishDirectory="PreserveNewest"`.
7. Unsupported OS/arch (e.g. macOS, x86): do nothing, no warning.

## Verify (all with the SDK at $SCRATCH/tools/dotnet/dotnet, from /home/user/QuickSheet)
Write any helper files ONLY under $SCRATCH/dlcheck/, never inside the repo (except obj/bin which are gitignored).
a) Empty pins (as committed): `dotnet build ExcelConsole.csproj` → succeeds, prints the "no pinned hash" message, 0 warnings, no libghostty file in bin/Debug/net9.0/.
b) Local server: create a dummy file `libghostty-vt-linux-x64.so.1` in the scratch dir, serve it with `python3 -m http.server 8765 --bind 127.0.0.1 --directory <scratch dir>` in the background, and build with `-p:LibGhosttyBaseUrl=http://127.0.0.1:8765/ -p:LibGhosttySha256_linux_x64=<its sha256 uppercase AND lowercase variants — check both work>` → file copied to bin/Debug/net9.0/libghostty-vt.so.1 and cached under obj/libghostty/...; second build doesn't re-download (check server log).
c) Wrong hash → warning about mismatch, nothing copied, nothing left in the cache; build still succeeds.
d) Server down → warning, build succeeds.
e) `-p:QuickSheetLibGhosttyDownload=false` → no download attempt, no warning.
f) Local override: dummy `native/linux-x64/libghostty-vt.so.1` takes precedence (no download even with a valid pin). DELETE native/ afterwards.
g) `dotnet publish -c Release -o <scratch>/pub` with a valid pin (local server) → libghostty-vt.so.1 present in publish dir.
h) `-getProperty` checks: with `-p:RuntimeIdentifier=linux-arm64`, QsNativeRid=linux-arm64 and QsNativeAsset=libghostty-vt-linux-arm64.so.1. Simulate Windows property computation as far as possible (you can't run Windows MSBuild; reason carefully and/or force properties).
Kill the http server at the end; clean bin/obj/obj/libghostty you created. Run `git -C /home/user/QuickSheet status --short` at the end and make sure only ExcelConsole.csproj and libghostty.props are yours.

Report back: the full final csproj native section + libghostty.props, the output of each verification (a–h), and concerns.
````
