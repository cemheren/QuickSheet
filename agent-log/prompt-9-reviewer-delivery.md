# Agent 9 — reviewer of dependency delivery (Opus)

Exact prompt (`$SCRATCH` = session scratch dir).

````text
You are a senior reviewer. Review the UNCOMMITTED working-tree changes in /home/user/QuickSheet (branch ccr-7ea14cd6-efxsg4): `git -C /home/user/QuickSheet diff` plus untracked files `.github/workflows/libghostty.yml`, `libghostty.props`, `scripts/libghostty-smoke/` (`git status --short`). DO NOT edit repo files or commit — review only. Read /home/user/QuickSheet/CLAUDE.md for context.

## Context
QuickSheet (.NET 9 desktop-wallpaper spreadsheet, installed from source via `git clone` → `dotnet run`) renders inline-command output with libghostty-vt (Ghostty's VT C library, release 1.0.0) via hand-written P/Invoke (Terminal/*.cs, already committed and reviewed). The maintainer decided to break the "zero dependencies" rule and take exactly one dependency on libghostty-vt so users don't need Zig. Ghostty publishes only source, so:
- `.github/workflows/libghostty.yml` builds it on 4 native runners (linux-x64, linux-arm64, win-x64, win-arm64) using scripts/build-libghostty.{sh,ps1} (official source + deps tarballs, pinned SHA-256, offline build), smoke-tests it via `scripts/libghostty-smoke` (compiles the real Terminal/*.cs), then the `publish` job creates an immutable GitHub Release `libghostty-vt-1.0.0-1` (never overwritten) and prints the 4 binary hashes.
- `libghostty.props` holds the release tag/URL and 4 SHA-256 pins (currently EMPTY — they'll be filled after the first CI run).
- `ExcelConsole.csproj` target `QsFetchLibGhostty` downloads the right asset into obj/, verifies the pin, and copies it to build/publish output; `native/<rid>/` local builds override; `-p:QuickSheetLibGhosttyDownload=false` opts out; failures never break the build (regex fallback at runtime).
- Docs updated for the "zero NuGet packages, one native dependency" policy.
IMPORTANT: Pushing this commit will trigger the workflow and publish the release (publish runs on any non-PR push while the release doesn't exist, deliberately, because pins must land before merge and workflow_dispatch requires the workflow on the default branch). A bug in what gets published is permanent under that tag (we'd have to bump the -N revision). So be thorough about anything that would make the published binaries wrong or the workflow insecure.

## Reference material (read-only)
- libghostty-vt 1.0.0 headers: $SCRATCH/vt100/include/ghostty/vt/
- Ghostty packaging guide (build flags, outputs, -Dtarget, Windows bin/ghostty-vt.dll): $SCRATCH/ghostty-build/PACKAGING-LIBGHOSTTY.md and build sources under .../ghostty-build/src/build/ (GhosttyLibVt.zig) — read only, do NOT run any ghostty build.
- Zig 0.17.0 official hashes: $SCRATCH/zig-index.json (key "0.17.0").
- .NET SDK: $SCRATCH/tools/dotnet/dotnet (build/run scratch stuff OUTSIDE the repo only; repo bin/obj are gitignored and OK).

## Focus
1. Workflow security: permissions scoping, script injection (`${{ }}` in run: blocks), which events can reach `publish` (forks, pull_request_target? tags?), GITHUB_TOKEN use, action pinning, whether a PR from a fork could ever publish or poison artifacts.
2. Workflow correctness: Zig URL/hash per platform (verify against zig-index.json), extraction dir names, PATH, `-Dtarget` with glibc suffix accepted by ghostty's build (check GhosttyLibVt.zig / build.zig use of standardTargetOptions), glibc check logic, Windows paths (`${{ runner.temp }}\out` in env + pwsh), the smoke-test env path vs where the scripts put the file (LIBGHOSTTY_OUT handling in both scripts), artifact names/paths, download-artifact layout (v4 puts each artifact in its own subdir), SHA256SUMS format, `gh release create` flags (`--latest=false`, `--prerelease=false` valid?), behavior if one matrix leg fails (publish skipped — acceptable?), immutability logic.
3. Build scripts (sh + ps1): the new ZIG_TARGET/LIBGHOSTTY_OUT handling, the ps1 bug fixes — PowerShell 5.1/7 correctness, `$LASTEXITCODE` handling, tar usage, hash comparisons; that checksum failure always aborts before building.
4. Smoke test (scripts/libghostty-smoke/Program.cs): are the expected strings actually what GhosttyVtScreen.GetText will produce on a real terminal (read Terminal/GhosttyVtScreen.cs: TrimStart/TrimEnd, unwrap, trim)? e.g. "red é\nprogress 99%\nkept" from the given byte stream on a 40x5 grid; scrollback test on 20x3 with 50 lines and maxScrollbackLines 100 (page-granular); after `\x1b[2J\x1b[H` — does the prior content move to scrollback (ghostty ED 2 behaviour) so GetText ends with "fresh"? A wrong expectation = failed CI = no release (safe but blocks). A too-weak check = bad binary published. Also: is the smoke csproj picked up by `dotnet build` at repo root or by quicksheet.sln?
5. csproj target: correctness of MSBuild logic (conditions evaluated at the right time, MSBuildLastTaskResult usage, cache re-verification, temp-file handling, case-insensitive compare), publish output, incremental build behaviour, that a missing/failed download can never break the build, Windows property computation (by reading), `BeforeTargets="AssignTargetPaths"` reliability, NormalizeDirectory with relative BaseIntermediateOutputPath, behaviour when building with `-r win-x64` on Linux.
6. Docs: accuracy vs the actual implementation (asset names, flags, what happens offline), and that no doc still claims "zero dependencies" for QuickSheet itself (grep README, CONTRIBUTING, CLAUDE.md, docs/, website/, .github/copilot-instructions.md, roadmap.md).

## Output
Findings most-severe first: severity (BLOCKER/MAJOR/MINOR/NIT), file:line, what's wrong, concrete failure scenario, specific fix. Only verified findings; mark uncertain ones PLAUSIBLE. Then a short "checked and correct" list. Concise.
````
