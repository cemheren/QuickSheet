# Agent 9 — reviewer of dependency delivery (Opus) — report

## Review: libghostty-vt release workflow, csproj download and docs (uncommitted changes)

**Verdict:** I found no blocker in workflow security. One issue should be fixed before pushing, because whatever gets published under `libghostty-vt-1.0.0-1` can't be replaced: the Windows ABI (finding 1). Everything else is minor.

What I ran: the MSBuild target in a scratch copy against a local HTTP server, covering good pin, tampered cache, wrong pin, 404, connection refused, stalled server, publish, `publish --no-build`, `-warnaserror` and RID evaluation. I also compiled the smoke project. I didn't edit any repo files.

### MAJOR

**1. `.github/workflows/libghostty.yml:55,62`: the Windows legs build with the MSVC ABI, not GNU.** `zig-target` is empty for Windows, so the build uses the native target.
- In the vt-1.0.0 source, `src/build/Config.zig` sets `query.abi = .msvc` whenever no ABI is given. This contradicts PACKAGING-LIBGHOSTTY.md ("gnu ABI, which is what Zig picks by default").
- Upstream CI never builds it this way. `.github/workflows/test.yml:430` uses `-Dtarget=x86_64-windows-gnu`, and the matrix at 863–865 says MSVC "doesn't work yet". Tier 1 support is "Windows x86_64 (gnu ABI)".
- Consequences:
  - The result depends on whichever Visual Studio and Windows SDK the runner image has.
  - The DLL's minimum OS version comes from the build host.
  - The DLL probably imports `VCRUNTIME140.dll` (dynamic link of msvcrt.lib + vcruntime.lib; see the comments at `GhosttyLib.zig:113-147`). That DLL isn't on a clean Windows install. I'm marking the exact import list PLAUSIBLE because I couldn't run a build.
- The smoke test can't catch this: the runners have VS and the VC++ redist installed. Affected users would silently get the regex fallback, and the binary is frozen under the tag.
- **Fix:**
  - Set `zig-target: x86_64-windows-gnu` and `aarch64-windows-gnu`. arm64 is tier 3, so let CI confirm it builds.
  - Add an import check after the build (`llvm-readobj --coff-imports` or `dumpbin /dependents`). Fail on anything outside KERNEL32/ntdll/ADVAPI32/`api-ms-win-*`/ucrtbase.
  - Optionally add a Linux equivalent: a `readelf -d | grep NEEDED` allowlist (libc, libm, ld-linux).

### MINOR

2. **`libghostty.yml:214`: the release notes are wrong.** They say glibc 2.28 covers "Ubuntu 18.04+", but 18.04 ships glibc 2.27. It should say Ubuntu 20.04+ (Debian 10 and RHEL 8 are correct).

3. **`libghostty.yml:221,212`: the title and notes are hard-coded.** The title says "QuickSheet build 1", and the notes hard-code 1.0.0 and the source SHA.
   - Bumping `RELEASE_TAG` to `-2`, as the header comment says to, would publish a wrong title.
   - `RELEASE_TAG` is also duplicated in `LibGhosttyReleaseTag`, with nothing checking the two match.
   - **Fix:** derive the title from `RELEASE_TAG`. When the release already exists, check out the repo and fail if the pins in `libghostty.props` or the tag differ from the release's `SHA256SUMS`.

4. **`libghostty.yml:9-14,182`: triggers and races.**
   - `push` has `paths` but no `branches`. GitHub doesn't apply path filters to tag pushes, so every tag push (e.g. app release tags) runs the 4-leg build. While the release is missing, any branch or tag push can publish it.
   - Two quick pushes race between `release view` and `release create`. The second fails safely but leaves a red run.
   - **Fix:** add `branches: ['**']` under `push`, put `if: ... && github.ref_type == 'branch'` on `publish`, and give `publish` a concurrency group (`concurrency: libghostty-publish`, `cancel-in-progress: false`).

5. **Actions are pinned by tag only: `actions/*@v4` at :73,145,173,190.** `download-artifact` runs in the job with `contents: write`, and upstream Ghostty pins by SHA. Pin by SHA. Also set `persist-credentials: false` on checkout, because the build step runs third-party `build.zig` code.

6. **No license or notice files ship with the binaries.** The DLL/.so links libghostty (MIT), simdutf, highway (Apache-2.0/BSD), uucode, wuffs and translate_c (`deps/libghostty-vt-third-party.json`), plus Zig's libc++ and compiler-rt. The release ships only binaries and `SHA256SUMS`, and the asset set is frozen once published. Add a `THIRD-PARTY-NOTICES` asset (or LICENSE plus that JSON) now.

7. **`ExcelConsole.csproj:110`: a stalled network blocks the build for 100 s.** `DownloadFile` uses HttpClient's 100 s timeout. Against a stalled server I measured 102 s per build (timeouts aren't retried). This repeats on every build until the cache exists, including login autostart (`dotnet run` in `docs/install-startup.md`). **Fix:** add `Timeout="20000"`; I checked that the parameter works.

8. **Docs still contradict the new policy.**
   - `SECURITY.md:23` still says supply-chain issues "are not a concern for the core". libghostty-vt now parses untrusted PTY bytes, so its vulnerabilities are in scope.
   - `.claude/skills/grow-quicksheet/SKILL.md:12` (tracked; it drives public posts) still says "zero-dependency".
   - `report.md` and `website/index.html:129` ("Every line is in the repo") are now inaccurate.
   - `CONTRIBUTING.md` bump instructions leave out the version and SHA pins in `scripts/build-libghostty.{sh,ps1}`.

### NIT

9. **`-warnaserror` turns an offline build into a failure** (verified: two errors, build FAILED). This contradicts "never break the build". You could suppress MSB3923 via `MSBuildWarningsAsMessages` and use a high-importance Message instead of a Warning.

10. **The csproj ignores the OS part of `-r`.** `-r linux-musl-arm64` resolves to `linux-arm64`, which pulls the glibc .so. `-r win-x64` on a Linux host resolves to `linux-x64` (both verified with `-getProperty`). Also an older issue: with no RID, the architecture comes from `OSArchitecture`. An x64 SDK on ARM64 Windows therefore copies the arm64 DLL into an x64 app. Using the arch from `$(NETCoreSdkRuntimeIdentifier)` fixes that. In all these cases the result is only a fallback, never a broken build.

11. **The smoke test could be stricter.** Assert `GhosttyNative.Version == "1.0.0"`. Also confirm it loaded the `QUICKSHEET_LIBGHOSTTY` path rather than a probed system copy.

12. **Smaller items:**
    - `Terminal/GhosttyNative.cs` header comment still describes build-script-only delivery.
    - The workflow has no `timeout-minutes`.
    - A failed `gh release create` can leave a draft. `release view` would then report it as existing, and the job fails on the missing `SHA256SUMS`, which needs manual cleanup.
    - Consider turning on the repo's Immutable Releases setting. Right now "immutable" holds only by convention; the pins are the real protection.

### Checked and correct
- **Zig downloads:** the URLs and SHA-256 for all 4 assets match `zig-index.json` 0.17.0. The extraction directory names (`zig-<arch>-<os>-0.17.0`), `GITHUB_PATH` handling, `-ine` hash comparison and `sha256sum -c` are correct.
- **Linux target flag:** `-Dtarget=<arch>-linux-gnu.2.28` goes through `standardTargetOptions`, which also takes `-Dcpu=baseline`. The Nix patch paths don't run with a baseline CPU.
- **Sanity step:** the glibc-floor logic is safe under pipefail with `|| true`. The `sort -V` comparison and `grep -w` on versioned symbols are correct.
- **Workflow security:**
  - `permissions` are read-only at top level, with write only on `publish`.
  - `${{ }}` appears only in `env:`/`with:`, never in `run:` text, so there's no script injection.
  - `pull_request` from forks can't publish, and artifacts are per-run, so fork PRs can't poison them.
  - The token is only used in `publish`.
- **Paths:** `LIBGHOSTTY_OUT` handling in both scripts matches the sanity, smoke and stage paths on Linux and Windows (`runner.temp` in step env; the Windows `\` path is absolute for `IsPathFullyQualified`).
- **Artifacts and release:** download-artifact v4's per-artifact subfolders are handled by `find`. The `SHA256SUMS` format and awk parsing are correct. `gh release create --prerelease=false --latest=false --target` are valid flags, and gh uploads assets while the release is still a draft. Publish being skipped when any leg fails is acceptable.
- **Scripts:**
  - Checksum failure always aborts before extraction and build in both scripts.
  - `Fail` exits with 1, and the `finally` cleanup still runs.
  - The `PROCESSOR_ARCHITEW6432` fallback and the inbox `System32\tar.exe` are both used.
  - The array-splatted zig arguments and `$LASTEXITCODE` checks are correct.
- **Smoke test:**
  - The project compiles against the real `Terminal/*.cs`.
  - Expectations match the 1.0.0 source:
    - The 40×5 stream gives `"red é\nprogress 99%\nkept"` after trim/unwrap and `GetText`'s Trim calls.
    - ED 2 (`.complete`) clears in place unless a semantic prompt was seen, so the text ends with "fresh".
    - The formatter's default content is the whole PageList from `getTopLeft(.screen)`, so scrollback is included and the 50-line test passes (it uses 1000 lines, not 100; either passes).
  - The chunk cut indices are all valid.
  - It isn't in `quicksheet.sln`. `ExcelConsole.csproj` drops `scripts/**` from Compile and None, and the smoke `bin/` and `obj/` are gitignored.
- **csproj target (all exercised):**
  - Conditions are evaluated when each task runs, and `MSBuildLastTaskResult` behaves correctly after a failed download.
  - A tampered cache is re-verified, deleted and re-downloaded.
  - Only a verified temp file reaches the cache, and a wrong pin is discarded and not copied.
  - 404 and connection-refused builds produce warnings and still succeed.
  - The pin is compared case-insensitively.
  - `NormalizeDirectory` resolves to `<proj>/obj/libghostty/<tag>/<rid>/`.
  - `BeforeTargets="AssignTargetPaths"` works for build, publish, `publish --no-build` and self-contained publish.
  - `native/<rid>` overrides the download, and empty pins skip it with a message.
- **README, CHANGELOG, CLAUDE.md, CONTRIBUTING, faq/tour and copilot-instructions:** these are accurate about asset names, the opt-out flag and offline fallback behaviour.
