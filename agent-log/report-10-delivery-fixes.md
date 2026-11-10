# Agent 10 — report

All requested fixes are in the working tree (nothing committed or pushed). The workflow and the Windows `.ps1` changes could not be run here. I verified them by YAML parse, `bash -n` on the extracted run blocks, and a local test of the `prop` sed helper only.

**Verification**
- **YAML and shell:** PyYAML parsed the workflow. `bash -n` passed on `build-libghostty.sh` and on all 8 extracted bash run blocks.
- **Main build:** `dotnet build ExcelConsole.csproj` gave 0 warnings and 0 errors with empty pins.
- **Smoke project:** it builds with 0 warnings and 0 errors. I could not run it, since there is no real libghostty binary here.
- **csproj scenarios (local `http.server` on 127.0.0.1:18765):**
  - A good pin downloaded the file and copied it to `bin/Debug/net9.0/libghostty-vt.so.1`.
  - A bad pin gave one SHA-256 mismatch warning, kept as a Warning.
  - Offline (closed port, pin set) with `-warnaserror` gave 0 warnings and 0 errors.
- **`-getProperty` checks (via `-p:RuntimeIdentifier=`):**
  - No RID: `linux-x64`.
  - `win-x64`: `win-x64`, `ghostty-vt.dll`, `ghostty-vt-win-x64.dll`.
  - `linux-arm64`: `linux-arm64`.
  - `linux-musl-arm64`, `osx-arm64` and `linux-arm`: empty.
- **`check_deps.py` ELF:** `libz.so.1` passes. dotnet's `libhostfxr.so` fails on `libstdc++.so.6` and `libgcc_s.so.1`, matching `readelf -d`.
- **`check_deps.py` PE:**
  - There is no native DLL in the container, so I used native PE executables from the uv cache. The parser read the PE32+ x64 and arm64 executables and the PE32 one, and it flagged `SHLWAPI.dll` correctly.
  - A managed .NET DLL reports `mscoree.dll` as disallowed. Another managed DLL with no imports gives an empty list.
  - A non-PE file is rejected with exit 1. The delay-load code path was not exercised, because no test file had one.
- **Cleanup:** the server is killed and its pid file removed. The repo's `bin/`, `obj/` and the smoke project's `bin/`/`obj/` are deleted and there is no `native/`. `git status` shows only expected files; the SECURITY.md, CONTRIBUTING.md, website and skill changes come from the other agent.

**Findings**
1. **Windows GNU ABI.** `.github/workflows/libghostty.yml` matrix now uses `zig-target: x86_64-windows-gnu` for win-x64 and `aarch64-windows-gnu` for win-arm64. `build-libghostty.ps1` already passes `-Dtarget=$env:ZIG_TARGET`.
   - New `scripts/libghostty-smoke/check_deps.py` parses PE import and delay-load tables and ELF DT_NEEDED entries in stdlib Python, with the allowlists you gave.
   - Workflow steps "Dependency allowlist (Linux)" and "(Windows)" run it after each build. The publish job runs it over all 4 artifacts before creating the release.
2. **Release notes.** They now say Debian 10+, Ubuntu 20.04+, RHEL 8+. I also added a line that Windows uses the GNU ABI and imports are allowlist-checked.
3. **Title and consistency:**
   - `LIBGHOSTTY_VERSION` and `LIBGHOSTTY_SRC_SHA256` are workflow-level env vars used by the notes.
   - The title is `libghostty-vt $LIBGHOSTTY_VERSION (QuickSheet build ${RELEASE_TAG##*-})`.
   - Publish fails if `LibGhosttyReleaseTag` in `libghostty.props` differs from `RELEASE_TAG`.
   - If the release already exists, it fails on a draft with "draft release left over from a failed publish; delete it and re-run". Otherwise it downloads `SHA256SUMS` and compares each non-empty pin case-insensitively.
4. **Triggers and races.** `push` has `branches: ['**']`. `publish` has `if: github.event_name != 'pull_request' && github.ref_type == 'branch'`, the `libghostty-publish` concurrency group, and a 15-minute timeout. Build legs have `timeout-minutes: 60`.
5. **Pinned actions**, all with a trailing tag comment. Every checkout has `persist-credentials: false`, including the new one in publish. `git ls-remote --tags` on `refs/tags/v4.*` returned these as the highest v4 tags (no peeled entries):
   - checkout `11d5960a326750d5838078e36cf38b85af677262` (v4.4.0)
   - setup-dotnet `67a3573c9a986a3f9c594539f4ab511d57bb3ce9` (v4.3.1)
   - upload-artifact `ea165f8d65b6e75b540449e92b4886f43607fa02` (v4.6.2)
   - download-artifact `d3f86a106a0bac45b974a628896c90dbdf5c8093` (v4.3.0)
6. **License notices.** Both build scripts copy `LICENSE` to `libghostty-vt-LICENSE.txt`, and `deps/libghostty-vt-third-party.json` to `libghostty-vt-third-party.json`. A missing LICENSE is an error; a missing JSON prints a note. I confirmed the JSON exists in the 1.0.0 source tree.
   - The linux-x64 leg checks that LICENSE is non-empty, then uploads the `notices` artifact with `if-no-files-found: warn`.
   - Publish attaches both files, with the JSON only if present, and includes them in `SHA256SUMS`.
7. **Download timeout.** `Timeout="20000"` is on `DownloadFile` in `ExcelConsole.csproj`.
9. **`-warnaserror` offline.** `MSB3923` is added to `MSBuildWarningsAsMessages`, and "Could not download" is now a high-importance Message. The SHA-256 mismatch stays a Warning.
10. **RID OS and arch.** The OS and arch logic in `ExcelConsole.csproj` is rewritten per your rules. The architecture comes from `NetCoreSdkRuntimeIdentifier`, with `OSArchitecture` as fallback when no RID is set. The local-override ItemGroup and the target conditions are unchanged, and the good-pin scenario above confirms they still work.
11. **Smoke test.**
    - `Terminal/GhosttyNative.cs` has `public static string? LoadedPath { get; private set; }`, set in all three `ProbeLoad` success paths, and the class header comment is updated as requested.
    - The smoke program asserts `Version == "1.0.0"`. It also asserts that the `Path.GetFullPath` of `LoadedPath` equals that of `QUICKSHEET_LIBGHOSTTY`, ignoring case on Windows.

**Deviations**
- I added `libghostty.props` to the workflow's push and pull_request `paths`, because the publish job now reads it.
- The `-getProperty` checks used `-p:RuntimeIdentifier=` rather than `-r`.
- In the `.ps1`, `LICENSE` and the JSON are read from `$SrcDir` (`deps\libghostty-vt-third-party.json`).
