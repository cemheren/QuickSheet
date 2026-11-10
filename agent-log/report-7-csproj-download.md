# Agent 7 — csproj pinned download (Sonnet) — report

Done. I changed only `/home/user/QuickSheet/ExcelConsole.csproj` and added `/home/user/QuickSheet/libghostty.props`. Checks a–g passed as specified on Linux. Windows (part of h) is checked by reading only. Nothing is committed. The http server is killed, and bin/obj/native are removed. `git status` shows other agents' files too (15 entries in total); mine are the csproj (M) and libghostty.props (??).

**libghostty.props** is exactly your contract. I only extended the comment: an empty hash means "not pinned yet", so the download is skipped.

**ExcelConsole.csproj**
- The top part is unchanged: TFMs and the `Compile Remove` / `None Remove` for native/** and scripts/**. The BOM is preserved.
- The native section is rewritten. Everything below is in that section, in order:
  - `<Import Project="libghostty.props" />`.
  - `QuickSheetLibGhosttyDownload` defaults to `true`.
  - `QsNativeArch` keeps the same logic as before (RID-derived, else host arch, lowercase).
  - `QsNativeSupportedArch=true` only for x64 and arm64.
  - On Windows (`IsOSPlatform('Windows')` plus a supported arch): `QsNativeRid=win-$(arch)`, `QsNativeFileName=ghostty-vt.dll`, `QsNativeAsset=ghostty-vt-win-$(arch).dll`.
  - On Linux: `linux-$(arch)`, `libghostty-vt.so.1`, `libghostty-vt-linux-$(arch).so.1`.
  - `QsNativeSha256` is picked with four explicit conditions, one per RID, because MSBuild has no indirect property lookup. My first attempt, `$(LibGhosttySha256_$(x))`, failed with MSB4184.
  - `QsNativeLocalPath` points at `native/<rid>/<file>`, and `QsNativeCacheDir` is `NormalizeDirectory(projdir, BaseIntermediateOutputPath, libghostty, tag, rid)`.
  - An ItemGroup adds the local override as a `None` item with Link, `CopyToOutputDirectory` and `CopyToPublishDirectory`, only if the file exists.
- The target `QsFetchLibGhostty` runs `BeforeTargets="AssignTargetPaths"`. That works empirically for both build and publish. Its condition is: RID set, download not `false`, not a DesignTimeBuild, no local override. It does the following:
  - If the pin is empty, it emits the high-importance "no pinned hash" message and does nothing else.
  - If a cached file exists, it hashes it and deletes it on mismatch.
  - If there is no cache, it downloads with `Retries=2` and `ContinueOnError=true` into `…/<rid>/tmp/<asset>`.
  - A download failure produces your warning text.
  - A downloaded file is hashed. On a match it is copied to the cache name. On a mismatch it produces the mismatch warning and is never copied. The tmp file is always deleted.
  - If the cache file exists at the end, it adds the `None` item with Link, `CopyToOutputDirectory`, `CopyToPublishDirectory` and `Visible=false`.
- Hash comparison is case-insensitive: both sides are uppercased.

**Deviation from your spec:** I used the built-in `GetFileHash` task and compared the results myself, instead of `VerifyFileHash`. `VerifyFileHash` with `ContinueOnError` would log its own extra warning on top of mine, which would give duplicate warnings.

**Verification**
- **a) Empty pins:** the build succeeded with 0 warnings and printed "libghostty-vt: no pinned hash for linux-x64 yet; skipping download…". Nothing ghostty-related was in `bin/Debug/net9.0/`.
- **b) Local server:**
  - With the lowercase pin, the download ran and `bin/Debug/net9.0/libghostty-vt.so.1` (300000 bytes) appeared, with 0 warnings. The cache held only `obj/libghostty/libghostty-vt-1.0.0-1/linux-x64/libghostty-vt.so.1`, with no tmp file left.
  - A second build with the uppercase pin did not download again (server log: 1 GET in total).
  - A fresh build with the uppercase pin downloaded again and copied the file, so both cases work (2 GETs in total).
- **c) Wrong hash (last hex digit changed):** one warning: "libghostty-vt: SHA-256 mismatch … (expected …A0, got …A6); the downloaded file was discarded and NOT used". The build succeeded, nothing was in the output, and the cache was empty.
- **d) Server down:** the build succeeded with 2 warnings. One is the task's own MSB3923 "Connection refused". The other is "Could not download libghostty-vt (…); inline commands will use basic ANSI stripping. Build offline with -p:QuickSheetLibGhosttyDownload=false to silence." Nothing was copied.
- **e) `-p:QuickSheetLibGhosttyDownload=false`:** 0 warnings, and the server GET count stayed at 3, so there was no download attempt.
- **f) Local override:** a dummy `native/linux-x64/libghostty-vt.so.1` containing "dummy" was copied to bin, and the server saw 0 GETs even with a valid pin. I deleted `native/` afterwards.
- **g) Publish:**
  - `dotnet publish -c Release -o <scratch>/pub` with a valid pin and the local server produced `pub/libghostty-vt.so.1` (300000 bytes), downloaded from the server.
  - Publish with only the local `native/` override also produced the file.
  - The grep for the publish summary lines printed nothing, but the files are the evidence.
- **h) `-getProperty`:**
  - With `RuntimeIdentifier=linux-arm64`: `QsNativeRid=linux-arm64`, `QsNativeAsset=libghostty-vt-linux-arm64.so.1`, `QsNativeFileName=libghostty-vt.so.1`, and an empty `QsNativeSha256`.
  - With `linux-x64` and `LibGhosttySha256_linux_x64=ABC`: `QsNativeSha256=ABC`.
  - With `linux-x86`: `QsNativeRid` is empty, so the target is skipped silently.
  - Windows could not be run. By reading, the Windows PropertyGroup uses the identical logic with `win-` prefixes, `ghostty-vt.dll` and `ghostty-vt-win-<arch>.dll`, and `QsNativeSha256` picks `LibGhosttySha256_win_x64` or `_win_arm64`.

**Concerns**
1. When a download fails you get two warnings (MSB3923 from `DownloadFile`, plus your message), because `ContinueOnError=true` converts the task's error into a warning.
2. The uppercased pin shows in the mismatch warning. That is cosmetic.
3. The GitHub release URL redirects to a CDN. `DownloadFile` follows redirects, but I could not test the real URL here.
4. A solution-wide `dotnet build` with a RuntimeIdentifier of a different OS (for example `win-x64` on Linux) leaves `QsNativeRid` as `linux-x64`, because the OS check wins. That matches the old behavior of selecting by host OS.
