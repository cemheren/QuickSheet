# AwesomeCSV submission — draft (2026-05-17)

> User submits the PR manually. Do not push to secretGeek/AwesomeCSV.

## Target

Repo: https://github.com/secretGeek/AwesomeCSV (926★, active, canonical CSV-tools list)
File: `README.md`
Section: **Tools** (where Tad, Modern CSV, csvkit, Miller, qsv live — interactive viewers/editors are mixed in here, not split out).
Contribution: maintainer references `contributing.md` — submitter should skim that before opening the PR for any list-specific rules.

## One-line entry (matches list's existing style)

The list's existing style: `[Tool Name](link) - Brief description, sentence-case.`

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive CSV grid that doubles as your desktop wallpaper. Cells run shell commands, sparklines, hyperlinks. Cross-platform (Windows WorkerW, Linux X11). MIT, .NET 9, zero dependencies.
```

Shorter variant (if the maintainer prefers terser entries):

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - CSV grid that runs as your desktop wallpaper. Cells can run shell commands. Windows + Linux, MIT.
```

Lead with **CSV** because that's the list's framing. Keep the wallpaper angle as the differentiator — it's what makes this row noticeably different from Tad and Modern CSV.

## PR title

```
Add QuickSheet (interactive CSV grid as desktop wallpaper)
```

## PR body

```markdown
Adds QuickSheet to the Tools section.

QuickSheet is an open-source (.NET 9, MIT, zero dependencies) CSV editor that runs in two modes:

- A normal terminal TUI in any shell — view/edit a CSV like sc-im / VisiData.
- A `--desktop` mode that embeds the same CSV grid as your **interactive wallpaper** (Win32 WorkerW on Windows, `_NET_WM_WINDOW_TYPE_DESKTOP` on X11).

The data file is plain CSV at rest — same data round-trips through Excel, vim, csvkit, etc. Cells can hold plain text or prefixed values (`r:` for runnable shell commands, `s:` for sparklines, URLs auto-detected, `ext:` to install extensions from a GitHub repo).

Differentiator vs Tad / Modern CSV (already on the list): those are read-mostly viewers; QuickSheet is an *interactive* grid where each cell can be a piece of working data, and the wallpaper-embedding makes it useful as an ambient view of CSV state.

**Repo:** https://github.com/cemheren/QuickSheet
**License:** MIT
**Section suggested:** Tools (alongside Tad, Modern CSV, csvkit).
```

## Notes for the submitter

- Existing list entries are alphabetical inside each section (loosely). QuickSheet inserts under "Q" — between any existing P and R entries.
- The maintainer (John Atten / secretGeek) is a long-time CSV evangelist; the bar is "is this genuinely a CSV tool" not "is it commercially polished." QuickSheet passes (file format is CSV, not a wrapper around something else).
- If asked "but the wallpaper part isn't a CSV feature" — yes, that's fine, the list includes plenty of tools whose primary use case is broader than CSV (Miller, qsv). The entry's CSV value is "edit a CSV grid that's always visible on your desktop." Lead with that.
- The maintainer has not merged a PR in a few months at time of writing — don't expect a fast turn-around; non-merge isn't a rejection signal.

## Why this list

- 926★, active, **canonical** in the CSV-tools niche.
- Inbound link from a list this size = real long-tail trickle (5–30 stars over a year, plus persistent search-result presence).
- QuickSheet is *genuinely* a CSV tool (CSV is the persistence format, by the project's own hard rule) — the entry is honest, not stretching.

## Ranking among awesome-* submissions

Per `research/awesome-list-fit.md`, all current ranked targets:

| List                                   | Star count | Hit-prob | Status        |
|----------------------------------------|------------|----------|----------------|
| luong-komorebi/Awesome-Linux-Software  | ~22k       | high     | draft `drafts/awesome-linux-software.md` |
| **secretGeek/AwesomeCSV**              | **926**    | **high** | **this draft** |
| 0PandaDEV/awesome-windows              | ~2.4k      | medium   | draft `drafts/awesome-windows.md` |
| awesome-csharp (uhub)                  | smaller    | high     | `drafts/awesome-csharp-windows.md` §1 |
| awesome-dotnet                          | medium     | high     | `drafts/awesome-lists.md` |
| awesome-tuis                            | medium     | marginal | `drafts/awesome-lists.md` (tone needs rewrite) |
| awesome-selfhosted                      | 293k       | nil      | DROPPED — wrong fit |

Recommended submission order if doing all four at once: Linux-Software → AwesomeCSV → awesome-csharp → awesome-windows. Stagger a few days between so the list ecosystem doesn't read it as a burst.
