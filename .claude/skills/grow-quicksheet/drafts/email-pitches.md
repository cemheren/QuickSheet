# Pitch emails: terminaltrove + console.dev

Drafted 2026-05-14, informed by `research/niche-communities.md`. Both are send-and-forget — single email each, no thread to babysit. Use your real email address.

---

## terminaltrove.com

**To:** `hello@terminaltrove.com`
**Subject:** `Tool suggestion: QuickSheet — terminal spreadsheet that doubles as the desktop wallpaper`

```
Hi,

I'd like to suggest QuickSheet for Terminal Trove. It's an interactive terminal spreadsheet that can also embed itself as the desktop wallpaper.

- Repo: https://github.com/cemheren/QuickSheet
- License: MIT, .NET 9
- Platforms: Windows + Linux (X11)
- Tour: https://github.com/cemheren/QuickSheet/blob/main/docs/tour.md

A few things that might fit your collection:
- Cell prefixes for runnable commands (`r:`), live subprocess output (`i:`), inline sparklines (`s:`), and extensions installed by git URL (`ext: github:user/repo`).
- Zero NuGet dependencies — Win32 WorkerW for the Windows wallpaper trick, raw X11 P/Invoke for the Linux side. Clone, build, run.
- Small ecosystem of standalone extension repos (weather, TLS cert checker, crypto/stock quotes, dictionary, etc.) that you can drop into a cell.

Happy to answer anything if useful, or to provide a different screenshot if there's a format you prefer.

Thanks,
Akif (https://github.com/cemheren)
```

Notes:
- Keep it < 200 words. Terminal Trove curators read a lot of these.
- Don't say "please feature us." Suggest, don't ask.
- The wallpaper-mode framing is the differentiator — every other suggestion they get is "yet another TUI."

---

## console.dev (weekly newsletter)

**To:** the contact-form email on `console.dev/contact` (currently obfuscated via Cloudflare; resolve at send-time).
**Subject:** `Suggestion for Console: QuickSheet (TUI spreadsheet, wallpaper mode, zero deps)`

```
Hi Console team,

Long-time reader. I want to suggest QuickSheet for a future issue — it lines up cleanly with the criteria on https://console.dev/selection-criteria.

QuickSheet is an interactive terminal spreadsheet for .NET 9 that can also embed itself as the desktop wallpaper. Repo: https://github.com/cemheren/QuickSheet.

How it scores against your criteria:

- Useful to developers: cell prefixes for runnable commands, live subprocess output, sparklines, hyperlinks. Wallpaper mode means it lives behind every window.
- Primary user is a developer: yes — TUI-first, CSV-only persistence, keyboard-driven.
- Self-service signup: open source, MIT, no sales contact.
- Regular-use tool: designed to be the always-on surface, not opened on demand.
- High quality / actively maintained: v0.2.0 tagged, CHANGELOG, CONTRIBUTING, SECURITY in place. Recent commits.
- Documentation: README + 60-second tour + recipes + extensions directory in docs/.
- Fast: no boot, no DB, opens an existing CSV instantly.
- Cross-platform: Windows + Linux.
- Advanced-user nods: CLI, keyboard shortcuts, headless `--export-md` for piping, theme presets.

Differentiator vs other TUI spreadsheets: the `--desktop` flag embeds the same grid as the wallpaper (Win32 WorkerW on Windows, `_NET_WM_WINDOW_TYPE_DESKTOP` on X11), so it's part of your environment rather than something you open. Zero NuGet dependencies — all the native interop (X11, WinForms, ConPTY) is hand-written P/Invoke.

Side project, but in active development with a public roadmap.

Thanks for considering,
Akif (https://github.com/cemheren)
```

Notes:
- Reference their published criteria explicitly — shows you read them, not pattern-matching the form.
- Mention v0.2.0 + the docs structure to signal "this is a maintained thing, not a one-off."
- Keep the differentiator in one sentence near the bottom. Their editorial style favors substance.

---

## Sending guidance

- Send terminaltrove first — they list more aggressively.
- Wait ~7 days before sending console.dev. If terminaltrove lists you and console.dev sees that, the second pitch is easier.
- Do not follow up unless you have new news (a release, a notable post, a milestone).
- If either responds, reply same-day with whatever they ask for (screenshot, additional info). Editorial inboxes have a short memory.

## What success looks like

- **terminaltrove**: listed in their "Newly Added" carousel within ~30 days = a steady trickle of stars from their site for months. Past "Tool of the Week" features have driven measurable spikes for similar projects.
- **console.dev**: a feature in the Thursday newsletter = a one-time spike of 50-300 stars depending on the issue's overall topic mix.
