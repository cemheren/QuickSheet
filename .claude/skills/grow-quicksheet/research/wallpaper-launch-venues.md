# Where wallpaper-mode tools actually launch (2026-05-17)

> Companion to `show-hn-tui-patterns.md`. That brief modelled QuickSheet as a TUI launch
> (peer = Bagels, 283 pts). After the README hero pivot (PR #108) names wallpaper-mode
> as the lead surface, the relevant peer set changes — and so does the venue.

## TL;DR (5 lines)

- **HN is the wrong primary venue for QuickSheet's wallpaper framing.** The wallpaper-tool genre consistently underperforms on HN: Übersicht (~5k★ peer) never broke 10 points across 5 HN submissions; most desktop-widget Show HNs score ≤6.
- **One exception**: "Parallax wallpaper engine for Linux and Windows" → **225 pts, 65 comments** (2023). That post lead with the *technical-curiosity* angle (animated 3D depth using DCT), not the productivity angle.
- The audience that consistently rewards wallpaper/desktop tools lives on **r/unixporn** (rice culture, screenshot-driven) and **r/selfhosted** (tool-utility, problem-solving).
- Recommended launch sequence: r/unixporn screenshot post FIRST → r/selfhosted "Homepage.io alternative not in a tab" SECOND → HN third, but reframe.
- For the HN reframe: lead with the **mechanism** (CSV + JSON-lines extension protocol + zero NuGet + cross-OS) not the wallpaper angle. HN responds to interesting technique, not "your wallpaper is X."

## Data — Show HN wallpaper-tool genre

Searched HN Algolia for "show hn wallpaper" + "show hn desktop dashboard" + Übersicht.

| Post | Points | Comments | Date | Notes |
|------|--------|----------|------|-------|
| Parallax wallpaper engine for Linux and Windows | **225** | 65 | 2023-02 | Outlier. Hooked on technique (depth maps, animated). github.com/jszczerbinsky/lwp |
| Übersicht: A Node.js based monitoring tool on your desktop | 9 | 4 | 2014-04 | Originally introduced Übersicht — never went viral on HN despite ~5k★ today |
| Übersicht, a geektool alternative for OS X | 4 | 0 | 2016-04 | Same tool re-pitched two years later |
| Übersicht – JS/CSS desktop widgets for OS X | 3 | 0 | 2014-10 | Third try |
| Crypto-currency trades with Übersicht | 3 | 0 | 2017-10 | Use-case angle |
| Übersicht (no descriptor) | 1 | 0 | 2017-05 | Bare-name submission |
| Dashboard – plugin-based desktop widget system for Linux | 2 | 1 | 2026-02 | Recent. Same generic-positioning failure mode. |
| Animated wallpaper – reacts to mouse | 4 | 0 | 2024-12 | Same |
| Wallpaper that grows as you ship | 2 | 2 | 2026-01 | Same |
| Automatic wallpaper changer for Gnome | 2 | 1 | 2020-09 | Same |

Search for "show hn conky rainmeter": **zero matching submissions in HN history**. Neither flagship project ever did a Show HN at all — they grew on r/unixporn (Rainmeter) and Linux distros (Conky).

## What this tells us

1. **HN's wallpaper/desktop-widget niche is dead-on-arrival.** Average <10 points. The one breakout (Parallax @ 225) led with technique, not "look at my dashboard." There's no audience there waiting to be activated.
2. The wallpaper category's actual cultural home is **rice culture (r/unixporn, individual dotfile repos, r/desktops)** for aesthetics, and **r/selfhosted + r/homelab** for utility.
3. HN can still work for QuickSheet — but it has to land as *"interesting technique"*, not *"interesting product."* The mechanism story (CSV-as-persistence + JSON-lines extension protocol + WorkerW/X11 P/Invoke + zero NuGet) is HN-shaped. The "your wallpaper is a spreadsheet" story is r/unixporn-shaped.

## Cross-reference with persona research

| Persona file | Top channel(s) named | Match with this data? |
|--------------|---------------------|------------------------|
| `personas/students.md` | r/unixporn (rice) | ✅ confirmed strongest |
| `personas/homelab.md` | r/selfhosted, selfh.st | ✅ confirmed |
| `personas/gamedev-ttrpg.md` | r/unixporn (GM-screen rice) | ✅ confirmed |
| `personas/developers-sre.md` | r/sre, r/devops + HN | ⚠ partial — HN works only with a mechanism-led pitch |
| `personas/traders.md` | r/battlestations, r/Daytrading | ✅ rice-adjacent for multi-monitor |

The persona research and this venue brief converge: **rice/selfhosted communities first, HN only with a reframed pitch later.**

## Implications (queue these)

### 1. Re-order the launch sequence (Bucket C — when ready)

Currently `drafts/showhn.md` is treated as the primary launch artifact. It should be **second-or-third in the sequence**, not first.

Recommended order:

1. **r/unixporn rice post.** Requires a real wallpaper screenshot — *this is the prerequisite the user has to produce*. Title like `[Linux/Hyprland] My Hyprland wallpaper is an interactive CSV spreadsheet — install in one cell`. The screenshot does 80% of the work.
2. **r/selfhosted "Homepage.io alternative, not in a tab"** post. Once the `health:` ext is live (currently scaffolded in `drafts/extensions/quicksheet-health-ext/`).
3. **HN — but with a rewritten body.** Lead with the protocol + zero-NuGet + WorkerW/X11 P/Invoke. The Bagels-style personal-itch opening from current `showhn.md` is still good; the *technical-curiosity hook in the body* should be sharpened.

Existing `drafts/showhn.md` already includes a "personal itch" opening per the prior TUI-launch brief — keep it, but add a paragraph specifically about the **JSON-lines extension protocol** as a technical hook (HN loves "open protocol that any language can implement").

### 2. Net-new draft (queue, not write this run): `drafts/unixporn-rice.md` revision

A `drafts/unixporn-rice.md` already exists per the drafts directory. Audit it next non-R run to make sure the title pattern matches what r/unixporn actually rewards in 2025–26 (typically `[distro/WM] description` format).

### 3. Don't open new HN-launch drafts beyond what already exists

The current `showhn.md`, `lobsters.md`, `twitter-thread.md`, `mastodon-bluesky.md`, `reddit-commandline.md`, `reddit-dotnet.md`, `reddit-programming.md`, `devto-blog.md`, `email-pitches.md` drafts collectively cover the venue space. **More drafts = filler.** New Bucket C output should *revise* existing drafts in light of new findings, not duplicate them.

## What this brief does NOT change

- `show-hn-tui-patterns.md` (the older brief) is still useful for *HN-specific tactics* (em-dash titles, personal-itch openings, comment defense). Keep it. Just don't treat HN as the primary venue.
