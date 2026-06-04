# awesome-dotnet + awesome-cli-apps submissions — draft (2026-06-03)

> User submits the PRs manually. Do not push to quozd/awesome-dotnet or agarrharr/awesome-cli-apps.

---

## 1. awesome-dotnet (21k ★)

**Repo:** https://github.com/quozd/awesome-dotnet  
**Target section:** CLI  
**PR template:** "Add QuickSheet to CLI section"

### Where in the file

Under `## CLI` — alphabetical order. Currently ends around `System.CommandLine`. QuickSheet would slot between "Spectre.Console" and "System.CommandLine" alphabetically (Q < S).

### Entry (one line, matches list style)

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive terminal spreadsheet that can embed itself as the desktop wallpaper. CSV-only persistence, cell prefixes for runnable commands, live subprocess output, sparklines, and 69+ extensions. Zero NuGet dependencies — all native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. Cross-platform (Windows + Linux).
```

### PR body

```markdown
## Add QuickSheet to CLI section

[QuickSheet](https://github.com/cemheren/QuickSheet) is a .NET 9 interactive spreadsheet that runs in the terminal or embeds as the desktop wallpaper.

**Why it fits the CLI section:**
- Primary interface is keyboard-driven TUI/desktop grid
- Headless CLI pipeline: `--export-md`, `--export-html`, `--export-json`, `--filter`, `--sort`, `--select`, `--head`, `--tail`, `--transpose`, `--info`, `--count`
- Usable in shell pipelines: `cat data.csv | quicksheet --filter "col1>10" --export-md`

**Key points:**
- MIT license
- .NET 9, zero NuGet dependencies (hard policy — supply chain minimalism)
- All platform interop (Win32 WorkerW, X11 P/Invoke, ConPTY) is hand-written
- Actively maintained with CHANGELOG, CONTRIBUTING, SECURITY
- 69+ extension repos (JSON-lines protocol over stdin/stdout)
- Cross-platform: Windows + Linux (X11)

**Differentiator vs existing entries:** Existing CLI tools in the list are libraries *for building* CLIs (CommandLineParser, Spectre.Console, etc.). QuickSheet is a CLI *application* — an actual end-user tool written in .NET. Closer in spirit to how `dotnet-outdated` or `dotnet-format` are listed.
```

### Submission checklist (from awesome-dotnet CONTRIBUTING.md)

- [ ] One link per PR
- [ ] Description starts with capital, ends with period
- [ ] Alphabetical order within section
- [ ] No trailing whitespace
- [ ] Check the project hasn't already been added

---

## 2. awesome-cli-apps (19k ★)

**Repo:** https://github.com/agarrharr/awesome-cli-apps  
**Target section:** Productivity > Spreadsheet  
**PR template:** "Add QuickSheet to Spreadsheet section"

### Where in the file

Under `## Productivity` → `### Spreadsheet`. Current entries: q, xsv, tabview, visidata, sc-im, csvkit, miller. QuickSheet slots alphabetically between `q` and `sc-im`.

### Entry (one line, matches list style)

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive terminal spreadsheet with desktop wallpaper mode, cell formulas, runnable commands, live subprocess output, sparklines, and 69+ extensions.
```

### PR body

```markdown
## Add QuickSheet to Spreadsheet section

[QuickSheet](https://github.com/cemheren/QuickSheet) is an interactive terminal spreadsheet that also embeds as the desktop wallpaper.

**What sets it apart from existing entries (visidata, sc-im, etc.):**
- **Desktop wallpaper mode** — `--desktop` flag embeds the grid behind all windows. Always visible, zero alt-tab. Win32 WorkerW on Windows, `_NET_WM_WINDOW_TYPE_DESKTOP` on X11.
- **Cell prefixes** — `r: cmd` runs shell commands, `i: cmd` shows live subprocess output in-cell, `s: {range}` renders sparklines, `ext: github:user/repo` installs extensions.
- **Extension ecosystem** — 69+ standalone repos that speak JSON-lines over stdin/stdout. Weather, stocks, TLS cert expiry, Docker status, dictionary, GitHub streak, etc.
- **Headless CLI pipeline** — `--export-md`, `--filter`, `--sort`, `--select`, `--head`, `--tail`, `--transpose` for use in shell pipelines.
- **Zero dependencies** — .NET 9, no NuGet packages. All native interop is raw P/Invoke.

CSV-only persistence. MIT license. Cross-platform (Windows + Linux). Actively maintained.
```

### Submission checklist (from awesome-cli-apps CONTRIBUTING.md)

- [ ] Link points to GitHub repo
- [ ] Description is concise (one sentence preferred)
- [ ] Correct section
- [ ] Alphabetical order
- [ ] No duplicates

---

## 3. Bonus: awesome-dotnet-core (if applicable)

**Repo:** https://github.com/thangchung/awesome-dotnet-core  
**Target section:** Tools (or Application Frameworks > CLI)

If this list has an active maintainer (check last merge date before submitting), use the same entry text as awesome-dotnet above. This list specifically targets .NET Core/5+/6+/7+/8+/9+ projects, so QuickSheet's .NET 9 targeting is a perfect fit.

---

## Sending strategy

1. **Submit awesome-cli-apps first.** It's the more natural fit (QuickSheet *is* a CLI spreadsheet app). The Spreadsheet section already has visidata and sc-im — QuickSheet is the same genre.
2. **Wait 1–2 weeks, then submit awesome-dotnet.** Having the awesome-cli-apps listing as social proof helps. The awesome-dotnet PR can reference it.
3. **awesome-dotnet-core** only if the repo shows activity in the last 6 months. Stale lists = wasted PR.

## What success looks like

- **awesome-cli-apps listing**: steady trickle of 1–5 stars/week from people browsing the Spreadsheet section. High-quality traffic — these are people actively looking for CLI spreadsheet tools.
- **awesome-dotnet listing**: .NET developer audience discovers QuickSheet. Many .NET devs don't browse TUI lists but do browse awesome-dotnet. Cross-pollination.
- Combined: a permanent backlink on two 19k+ star repos → SEO + discovery flywheel.
