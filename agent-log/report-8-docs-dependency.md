# Agent 8 — docs for single-dependency policy (Haiku) — report

Docs update for the libghostty-vt policy change is done. Nothing committed. I touched only my owned files.

Files I changed (7 of them, all mine):
- README.md
- CONTRIBUTING.md
- CLAUDE.md
- CHANGELOG.md
- docs/faq.md
- docs/tour.md
- website/index.html
- .github/copilot-instructions.md (it states the dependency rule)

Note on `git diff --stat`: it also lists ExcelConsole.csproj, scripts/build-libghostty.ps1, scripts/build-libghostty.sh and an untracked libghostty.props. Those are from the other agents and I did not touch them. My own files add up to about 8 files (the README badge counts as an extra change inside README.md).

New text of each changed passage:

README.md
- Badge: `![Zero NuGet packages](https://img.shields.io/badge/NuGet_packages-0-brightgreen)` (the old "dependencies-0" badge would now be wrong).
- Bullet: "**One dependency.** Clone → `dotnet build` → run. No NuGet packages, no npm, no Docker. The only dependency is libghostty-vt, built from Ghostty's pinned source by our CI and hash-verified on download."
- Section renamed "### Terminal rendering (libghostty-vt)": "Inline `i:` command output is rendered by libghostty-vt … `dotnet build` downloads a prebuilt copy for your platform (linux-x64, linux-arm64, win-x64, win-arm64) from this repo's GitHub Releases and verifies its SHA-256 against hashes pinned in `libghostty.props`. No Zig or other tools are needed." Then three bullets: "**Opt out** … `-p:QuickSheetLibGhosttyDownload=false`", "**Build it yourself** (Zig 0.17 required) … `scripts/build-libghostty.sh` / `powershell … build-libghostty.ps1`. A library in `native/<rid>/` takes precedence over the download.", "**Point at a specific library:** set `QUICKSHEET_LIBGHOSTTY` to a full path." Closing line: "If the library is missing or fails to load (offline first build, macOS, unsupported architecture), QuickSheet falls back to basic ANSI stripping and prints a one-line note to stderr."
- "A note on the code": "The hard rules the project keeps are **zero NuGet packages** and **one native dependency**. All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. The only external dependency is **libghostty-vt** … CI builds it from Ghostty's pinned source, and `dotnet build` downloads it with a verified SHA-256. Without it QuickSheet falls back to basic ANSI stripping. Clone, build, run."

CONTRIBUTING.md
- Intro: "QuickSheet is a .NET 9 spreadsheet with zero NuGet packages and one native dependency (libghostty-vt) that runs in the terminal …"
- Row: "**Zero NuGet packages** | Supply-chain safety. … The only external dependency is libghostty-vt. Any new dependency needs an issue first — there's almost always a stdlib path."
- Row: "**libghostty-vt is the one native dependency** | CI builds it from Ghostty's official, SHA-256-pinned source tarball and publishes the binaries as a GitHub Release. `dotnet build` downloads them and verifies the hashes pinned in `libghostty.props`. To bump the version, update the workflow and `libghostty.props` together. Never commit native binaries (`native/` is gitignored), and the app must keep working when the library is absent (ANSI-stripping fallback)."

CLAUDE.md
- Build line: "./scripts/build-libghostty.sh   # optional (dotnet build downloads a prebuilt copy): build libghostty-vt (Zig 0.17) into native/ (Linux; .ps1 on Windows)"
- Conventions: "**Zero NuGet packages; one native dependency.** libghostty-vt is the only external dependency (see `libghostty.props`, `.github/workflows/libghostty.yml`). Do not add packages or other dependencies."
- Conventions: "**libghostty-vt delivery.** CI (`.github/workflows/libghostty.yml`) builds it from Ghostty's pinned, SHA-256-verified source and publishes GitHub Release binaries. The csproj downloads the right one and verifies it against hashes in `libghostty.props` (skip with `-p:QuickSheetLibGhosttyDownload=false`). A library in `native/<rid>/` overrides the download. Never commit native binaries, and the app must still run without it (`LegacyVtScreen` fallback)."

CHANGELOG.md (under "## Unreleased")
- Changed entry now ends: "libghostty-vt is downloaded automatically by `dotnet build`: CI builds it from Ghostty's pinned source with a SHA-256-verified download, so no Zig is needed. Without it QuickSheet falls back to the old stripper."
- New heading "### Dependencies": "- QuickSheet now has one external dependency, libghostty-vt (still no NuGet packages)."

docs/faq.md, "Why .NET?"
- "The hard rules are **zero NuGet packages** and **one native dependency**. All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. The one dependency is **libghostty-vt** … CI builds it from Ghostty's pinned source, and `dotnet build` downloads it with a verified SHA-256. No Zig is needed. Without it QuickSheet falls back to basic ANSI stripping."

docs/tour.md, "6. The hard rules"
- "**Zero NuGet packages, one native dependency.** All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. Clone, build, run. The only external dependency is **libghostty-vt** … CI builds it from Ghostty's pinned source, and `dotnet build` downloads it with a verified SHA-256. Without it QuickSheet falls back to basic ANSI stripping."

website/index.html, trust list
- "<li><strong>Zero NuGet packages.</strong> Every line is in the repo. Native interop is hand-written P/Invoke (libX11, Win32). The one native dependency, libghostty-vt, is built from Ghostty's pinned source by our CI and hash-verified on download.</li>"
- Left the meta line "MIT-licensed · Zero NuGet deps · …" alone. It is still true.

.github/copilot-instructions.md, "Key Conventions"
- "**Zero NuGet packages; one native dependency** — Intentional policy to avoid supply chain risk. All platform interop is done via raw P/Invoke. The only external dependency is libghostty-vt, downloaded by `dotnet build` with a pinned SHA-256 (see `libghostty.props`). Do not add NuGet packages or other dependencies."

Deliberately left alone:
- docs/extensions.md, docs/extension-protocol.md, docs/for-homelab.md, and the "zero NuGet" wording in docs/extensions.md rows. These describe the rules for extension repos, not QuickSheet itself, so they remain accurate.
- README.md "Zero cloud" and "zero network" phrasing, and CHANGELOG historical entries (lines 117, 310 and older). These are history or unrelated.
- roadmap.md, report.md, SECURITY.md, .gitignore comments and the .claude/ and .agents/ grow-quicksheet drafts. Not in my owned list. roadmap.md:254 still says "zero NuGet" as a hard policy and is worth a look by its owner.
- Also not owned, but noted: .gitignore:13 already matches the new policy.
