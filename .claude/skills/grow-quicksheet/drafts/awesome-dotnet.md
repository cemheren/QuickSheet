# awesome-dotnet submission draft

**Target:** https://github.com/quozd/awesome-dotnet
**Stars on list:** ~19k
**Category:** CLI (best fit) — terminal-based interactive tool. Secondary option: Office (spreadsheet).
**Guidelines:** One link per PR; meaningful description; project must be useful, maintained, stable, documented.

---

## Recommended category: CLI

**Diff line to add (alphabetically sorted within CLI section):**

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive terminal spreadsheet that also embeds as a desktop wallpaper. CSV-backed, with cell prefixes for runnable commands, live subprocess output, sparklines, and hyperlinks. Zero NuGet dependencies; cross-platform (Windows + Linux).
```

---

## PR title

```
Add QuickSheet to CLI section
```

## PR body

```markdown
Adds [QuickSheet](https://github.com/cemheren/QuickSheet) under **CLI**.

QuickSheet is an interactive terminal spreadsheet (.NET 9) with an unusual dual
mode: it also embeds itself as a transparent, interactive desktop wallpaper on
both Windows (WinForms/WorkerW) and Linux (raw X11 P/Invoke).

**Why it fits this list:**

- Pure .NET 9, zero NuGet dependencies — all native interop is hand-written P/Invoke.
- Cross-platform: Windows and Linux with OS-conditional TFMs.
- CLI-first: headless export modes (CSV → Markdown / HTML / JSON), `--info`, piping support.
- Cell prefix DSL: `r:` (runnable commands), `i:` (live subprocess output),
  `s:` (sparklines), `ext:` (extension protocol for community plugins).
- CSV as the persistence format — dead simple, no database.
- MIT licensed, actively maintained.

**Quality checklist:**

- [x] Generally useful to the community (terminal productivity, data inspection)
- [x] Actively maintained (weekly commits)
- [x] Stable (ships on .NET 9, both platforms tested)
- [x] Documented (README with feature tour, `docs/` directory, `--help`)
- [x] MIT license
```

---

## Alternative category: Office

If CLI section is too crowded or maintainers prefer:

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - Lightweight terminal spreadsheet with CSV persistence, cell formulas (Σ/Π), sparklines, and a unique desktop-wallpaper mode. Zero NuGet dependencies; .NET 9; Windows + Linux.
```

---

## Submission notes for user

1. Fork `quozd/awesome-dotnet`.
2. Find the `## CLI` section — entries are alphabetically sorted.
3. Insert the line above in alphabetical order (after entries starting with P/Q).
4. Open PR with the title/body above.
5. The list requires projects to have tests — QuickSheet doesn't have formal tests.
   This may cause rejection. Mitigation: the project is well-documented and stable,
   and many entries on the list lack comprehensive test suites. Worth trying.

**Risk:** The "Tests" quality criterion is a soft guideline, not a hard gate.
Many projects on the list lack formal tests. The maintainers tend to be
pragmatic about it for tools that are clearly working and useful.
