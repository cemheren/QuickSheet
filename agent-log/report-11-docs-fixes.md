# Agent 11 — remaining policy docs (Haiku) — report

Docs update for libghostty-vt policy is done. Nothing committed. `git diff --stat -- SECURITY.md CONTRIBUTING.md website/index.html .claude/` shows 15 files changed (31 insertions, 21 deletions). No files outside my scope were touched.

**1. SECURITY.md, "## Dependencies" (replaced the old paragraph)**
New text:
"QuickSheet's core has **zero NuGet packages** as a deliberate policy. Its one native dependency is [libghostty-vt](https://github.com/ghostty-org/ghostty), which renders inline `i:` output and parses the PTY bytes those commands produce, so it is in scope. Vulnerabilities in libghostty-vt itself are in scope, and so are flaws in how QuickSheet builds, downloads, verifies or loads it: the `.github/workflows/libghostty.yml` build from Ghostty's pinned source, the SHA-256 pins in `libghostty.props`, and `QUICKSHEET_LIBGHOSTTY` loading. Other third-party packages are not a concern for the core. Reference extensions follow the same NuGet rule."
I kept the existing In/Out-of-scope bullets unchanged. The section heading and structure are the same.

**2. CONTRIBUTING.md**
- The libghostty-vt table row now reads: "...To bump it, see [Updating libghostty-vt](#updating-libghostty-vt). Never commit..." The rest of the row is unchanged.
- New subsection "### Updating libghostty-vt" added between the ground-rules table and "## What's helpful":
  1. Set the new version and its source/deps SHA-256 pins in `scripts/build-libghostty.sh` and `scripts/build-libghostty.ps1`.
  2. Set `RELEASE_TAG` and the version/source-SHA env vars in `.github/workflows/libghostty.yml`.
  3. Set `LibGhosttyReleaseTag` in `libghostty.props`.
  4. Push. The workflow publishes an immutable GitHub Release only if the tag doesn't exist yet. To rebuild the same upstream version, bump the trailing `-N` of the tag (e.g. `libghostty-vt-1.0.0-2`).
  5. Paste the four hashes from the workflow run summary into `libghostty.props` and push that too.
- The "Zero NuGet packages" row and the intro line were already accurate and are unchanged.

**3. website/index.html, line 129**
Old: "Every line is in the repo."
New: "All of QuickSheet's code is in the repo; native interop is hand-written P/Invoke (`libX11`, Win32). The one native dependency, libghostty-vt, is built from Ghostty's pinned source by our CI and hash-verified on download."
Other claims checked: the meta description (line 7) and og:description (line 9) do not mention dependencies. The badge text on line 36 "Zero NuGet deps" stays, since it is still true. No JSON-LD or other "zero dependencies" claims exist on the page.

**4. .claude/skills/grow-quicksheet/SKILL.md, line 12-13 (mission)**
Old: "The project is a zero-dependency / .NET terminal spreadsheet"
New: "The project is a zero-NuGet / .NET terminal spreadsheet (one native dependency: libghostty-vt)"

Other changes in that folder, each a phrase swap about QuickSheet itself:
- drafts/awesome-csv.md: "MIT, .NET 9, zero dependencies." becomes "MIT, .NET 9, zero NuGet packages." Also "(.NET 9, MIT, zero dependencies)" becomes "(.NET 9, MIT, zero NuGet packages)".
- drafts/awesome-linux-software.md: "MIT, zero deps, X11" becomes "MIT, zero NuGet packages, X11".
- drafts/email-pitches.md: subject "wallpaper mode, zero deps)" becomes "wallpaper mode, zero NuGet)".
- drafts/reddit-dotnet.md: "zero-dependency code" becomes "zero-NuGet code". "zero-deps rule" becomes "zero-NuGet rule" (2 places). "zero-deps claim" becomes "zero-NuGet claim".
- drafts/reddit-commandline.md and drafts/showhn.md: "zero-deps rule" becomes "zero-NuGet rule".
- research/awesome-list-fit.md: "MIT, zero deps." becomes "MIT, zero NuGet packages."
- research/personas/students.md, gamedev-ttrpg.md: "(zero deps, MIT)" becomes "(zero NuGet, MIT)".
- research/personas/developers-sre.md: "(zero deps, open source)" becomes "(zero NuGet, open source)".
- research/personas/homelab.md: "(zero deps, MIT, extension protocol)" becomes "(zero NuGet, MIT, extension protocol)".

**Deliberately left alone**
- Every "zero NuGet" / "zero NuGet deps" / "NuGet deps added" phrase in the skill folder and the website. These claims are still true, since QuickSheet still has no NuGet packages.
- The "zero-dep" phrasing in SKILL.md line 183 ("any zero-dep language"), in drafts/twitter-thread.md:51 and drafts/mastodon-bluesky.md:65 ("all zero-dep, all MIT"), and in research/accounting-extensions.md:45 ("Zero deps."). All of these describe extensions or the language a new extension may use, not QuickSheet core.
- All of drafts/extensions/** (extension READMEs saying "zero NuGet dependencies") and the extension notes in research/. These describe separate repos, and "zero NuGet" is still accurate for them.
- log.md. It is a historical action log, and its entries describe past states.
- README.md and CLAUDE.md, which are outside my scope. Both already say "zero NuGet packages / one native dependency" correctly. CLAUDE.md line 42 is the source of truth for this policy.

I did not verify the links or anchors by rendering. The anchor `#updating-libghostty-vt` matches the new heading under GitHub's slug rules.
