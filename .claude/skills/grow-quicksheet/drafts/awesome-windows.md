# awesome-windows submission — draft (2026-05-17)

> User submits the PR manually. Do not push to 0PandaDEV/awesome-windows.

## Target

Repo: https://github.com/0PandaDEV/awesome-windows (2.4k★, active 2024–2026)
File: `README.md`
Section: **Customization** (primary) or **Productivity** (fallback).
License-aware: maintainer rejects "vibecoded slop" — keep the entry sober and concrete.

> NOTE: An older draft (`awesome-csharp-windows.md`) targeted `Awesome-Windows/Awesome`,
> which now 404s on GitHub. That submission slot is dead. Use this one instead.

## One-line entry (matches list's existing style)

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) `🔓` `🆓` — Transparent, interactive spreadsheet that lives behind every window via the WorkerW desktop trick. Cells can run shell commands, stream subprocess output, render sparklines, and install JSON-lines extensions in one cell. .NET 9, zero NuGet dependencies.
```

If the list uses different badge tokens, mirror whatever the maintainer's current convention is.

## PR title

```
Add QuickSheet (Customization)
```

## PR body

```markdown
Adds QuickSheet to the Customization section.

**What it is:** an open-source desktop app that replaces your Windows wallpaper with a transparent, interactive spreadsheet grid. The grid sits *behind* every window — click anywhere on the desktop to take a quick note, paste a URL to make it clickable, prefix a cell with `r: code .` to make it a launcher, or install a JSON-lines extension with `ext: github:user/repo`.

**Why it fits Customization:** uses the standard Windows `WorkerW` desktop-painting trick (same family of techniques as Rainmeter/Stardock Fences) but with a CSV grid as the surface instead of pre-baked widgets. There's nothing else like it on the list.

**Repo:** https://github.com/cemheren/QuickSheet
**License:** MIT (`🔓 🆓`)
**Maintenance:** active, 40+ companion extensions, .NET 9.

**Truthful caveats:**
- It's Windows + Linux (Linux uses X11; Wayland warns and falls back to TUI).
- 1 star at time of submission — it's small. Mention if relevant to the maintainer's quality bar.
```

## Notes for the submitter

- The maintainer is explicit ("Vibecoded slop... will be rejected"). Lead with the WorkerW technique, not the marketing pitch. Concrete > breathless.
- Open the PR with a single, clean commit (no rebase noise, no merge bubbles). The list's recent merged PRs are all one-commit single-line additions.
- If asked "why not just Rainmeter?" the answer is: Rainmeter renders pre-built skins; QuickSheet renders editable cells you type into. Different primitive.
- If rejected on quality bar: don't argue. Accept silently. The list is theirs.

## After submission

- Inbound link from a 2.4k-star list. Discovery via search results + the list's own audience checking new additions. Realistic: 5–20 extra stars in long-tail traffic if accepted.
- The list is mirrored / featured on awesome-* aggregators, so an accepted entry compounds.

## Ranking among awesome-* submissions

Per `research/awesome-list-fit.md`:

| List                                   | Star count | Hit-prob | Status        |
|----------------------------------------|------------|----------|----------------|
| luong-komorebi/Awesome-Linux-Software  | ~22k       | high     | draft ready (drafts/awesome-linux-software.md) |
| **0PandaDEV/awesome-windows**          | ~2.4k      | medium   | **this draft**  |
| awesome-csharp                          | smaller    | high     | drafts/awesome-lists.md |
| awesome-dotnet                          | medium     | high     | drafts/awesome-lists.md |
| awesome-tuis                            | medium     | marginal | drafts/awesome-lists.md (tone needs rewrite to lead with wallpaper) |
| awesome-selfhosted                      | 293k       | nil      | DROPPED — wrong fit |
