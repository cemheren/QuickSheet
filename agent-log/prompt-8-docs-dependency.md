# Agent 8 — docs for single-dependency policy (Haiku)

Exact prompt (`$SCRATCH` = session scratch dir).

````text
You are editing documentation in /home/user/QuickSheet (C#/.NET 9 desktop-wallpaper spreadsheet). Branch ccr-7ea14cd6-efxsg4. Do NOT commit. Other agents are editing ExcelConsole.csproj, libghostty.props, .github/workflows/, scripts/ in parallel — do NOT touch those. You own ONLY: README.md, CONTRIBUTING.md, CLAUDE.md, CHANGELOG.md, docs/faq.md, docs/tour.md, .github/copilot-instructions.md (only if it mentions the dependency rule), and any other file under docs/ that states "zero dependencies"/"no NuGet" (grep for: `zero NuGet`, `Zero NuGet`, `zero dependencies`, `Zero dependencies`, `No NuGet`, `no NuGet`, `supply chain`, `libghostty`).

## The policy change (decided by the maintainer)
QuickSheet is breaking its "zero dependencies" rule to take exactly ONE external dependency: **libghostty-vt**, the terminal-emulator library from the Ghostty terminal, used to render inline `i:` command output (cursor movement, `\r` progress bars, clear-screen, full-screen TUIs like htop show their real final screen + scrollback).

How it's delivered (describe accurately, concisely):
- Still **no NuGet packages** (and none should be added) — the dependency is a native library.
- Ghostty publishes only source, so QuickSheet's own GitHub Actions workflow (`.github/workflows/libghostty.yml`) builds libghostty-vt from Ghostty's **official, SHA-256-pinned source tarball** (release 1.0.0), smoke-tests it through QuickSheet's P/Invoke binding on linux-x64, linux-arm64, win-x64 and win-arm64, and publishes the binaries as an immutable GitHub Release of this repo (`libghostty-vt-1.0.0-1`).
- `dotnet build` **downloads the right binary automatically** and verifies it against SHA-256 hashes pinned in `libghostty.props`. Users need nothing extra — still clone → `dotnet run`. No Zig.
- Offline/opt-out: `-p:QuickSheetLibGhosttyDownload=false`. If the library is missing or fails to load (offline first build, macOS, unsupported arch), QuickSheet falls back to basic ANSI stripping and prints one line to stderr.
- Contributors who want to build it themselves can still run `scripts/build-libghostty.sh` / `powershell -ExecutionPolicy Bypass -File scripts/build-libghostty.ps1` (needs Zig 0.17); a library in `native/<rid>/` takes precedence over the download. `QUICKSHEET_LIBGHOSTTY=<absolute path>` overrides at runtime.
- Updating it: bump the version/tag in the workflow + `libghostty.props`, run the workflow, paste the four hashes from its run summary into `libghostty.props`.

## Edits
Read each file first; keep each project's voice; be concise (don't add marketing). The previous commit already added libghostty text describing it as an "optional" library built by the scripts with Zig — REWRITE those passages so they match the new reality (automatic download; scripts are only for contributors who want to build it themselves). Specifically:
- README.md:
  - The "Zero dependencies" bullet near the top (`Clone → dotnet build → run. No NuGet packages, no npm, no Docker. The entire supply chain is the .NET SDK.`) → becomes a "One dependency" bullet: still clone → build → run, no NuGet/npm/Docker; the only dependency is libghostty-vt, built from Ghostty's pinned source by our CI and hash-verified on download.
  - Replace the "### Optional: libghostty-vt" subsection with a short "### Terminal rendering (libghostty-vt)" subsection: downloaded automatically and verified on first build; how to opt out; how to build it yourself; the fallback.
  - The "A note on the code" paragraph: update the rule statement (zero NuGet packages; one native dependency, libghostty-vt, handled as above).
- CONTRIBUTING.md: the ground-rules table rows — "Zero NuGet dependencies" stays (still true: no packages) but clarify that the only external dependency is libghostty-vt and new dependencies need an issue first; rewrite the "Native libraries are optional and built, not committed" row → "libghostty-vt is the one native dependency": built by CI from pinned official source, downloaded with pinned SHA-256, never commit binaries, app must keep working (fallback) when it's absent; how to bump it.
- CLAUDE.md: Conventions — change "**Zero NuGet deps.** Hard policy…" to "**Zero NuGet packages; one native dependency.** libghostty-vt is the only external dependency (see `libghostty.props`, `.github/workflows/libghostty.yml`). Do not add packages or other dependencies." Update the "Native libs are optional" bullet to describe: CI-built from pinned source, downloaded + SHA-256-verified by the csproj (`QuickSheetLibGhosttyDownload=false` to skip), `native/<rid>/` local builds override, never commit binaries, app must still run without it. In Build & Run keep the build-libghostty line but mark it "(optional — dotnet build downloads a prebuilt copy)". Keep the rest.
- CHANGELOG.md: in the existing "## Unreleased" section, update the existing libghostty entry so it says the library is downloaded automatically (CI-built from pinned source, hash-verified) — no Zig needed — and add a line under a "### Dependencies" heading: "QuickSheet now has one external dependency, libghostty-vt (still no NuGet packages)."
- docs/faq.md "Why .NET?" paragraph and docs/tour.md "The hard rules" bullet: update the zero-dependency statement accordingly (one or two sentences).
- Anything else your grep finds stating zero dependencies (e.g. website/ folder? — check `website/` too; if it has HTML/MD claiming zero dependencies, update the claim minimally).

Finally run `git -C /home/user/QuickSheet diff --stat` and report: every file you changed and the new text of each changed passage (quote them), plus anything you found but deliberately left alone.
````
