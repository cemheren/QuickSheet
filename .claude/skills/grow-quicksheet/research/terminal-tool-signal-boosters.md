# Terminal Tool Signal-Boosters & Influencers

**TL;DR:** 10 accounts/platforms that amplify TUI/CLI projects. QuickSheet's best
shots are: (1) TerminalTrove submission, (2) Console.dev submission, (3) being
visible to Will McGugan and Charm team on Twitter/X, (4) awesome-tuis PR via
rothgar. Each has a clear submission path documented below.

---

## Tier 1 — Curated Directories (submit directly, highest ROI)

### 1. TerminalTrove (@terminaltrove)
- **What:** The largest living directory of terminal tools. "Tool of the Week" feature.
- **Platform:** Website + Mastodon (@terminaltrove@mastodon.social) + Twitter/X (@terminaltrove).
- **How to submit:** https://terminaltrove.com/post/ — form requires: name, tagline (~100 chars), description (250–300 chars), 2–3 features, preview image (PNG/GIF), primary language, license, categories.
- **Criteria:** Cross-platform preferred, standalone binaries preferred, not a duplicate.
- **Alt:** Email curator@terminaltrove.com.
- **Fit for QuickSheet:** Strong. It's cross-platform, unique (spreadsheet-as-wallpaper), has a clear category (productivity / TUI / spreadsheet).

### 2. Console.dev (@consoledotdev)
- **What:** Curated weekly newsletter for experienced developers. High editorial bar.
- **Editors:** David Mytton, Jon Bromley.
- **How to submit:** https://console.dev/submit/ — name, description, why it matters, repo link.
- **Also:** Tag @consoledotdev on Twitter/X, DM.
- **What gets featured:** Developer productivity, rich TUI, good docs/onboarding, unique use case, open-source, easy install.
- **Fit for QuickSheet:** Good if positioned as "desktop productivity meets terminal culture." The zero-dep + wallpaper angle is novel enough for their editorial taste.

### 3. awesome-tuis (GitHub, @rothgar)
- **What:** 200+ item curated list of TUI projects. 7k+ stars.
- **Maintainer:** rothgar (GitHub handle). Accepts PRs with new entries.
- **How:** PR to https://github.com/rothgar/awesome-tuis with one-line entry.
- **Fit:** QuickSheet is literally a TUI spreadsheet. Slam dunk. Draft already exists in `drafts/awesome-lists.md`.

---

## Tier 2 — Individual Signal-Boosters (be visible to, don't spam)

### 4. Will McGugan (@willmcgugan)
- **Known for:** Creator of Rich and Textual (Python TUI libraries). Actively retweets cool TUI projects — even ones not built with his tools.
- **Platform:** Twitter/X, very active. Mastodon also.
- **Strategy:** When QuickSheet is posted to HN/Reddit and gets traction, a mention or reply that reaches Will's feed can snowball. He's enthusiastic about anything that pushes the "terminal is beautiful" narrative.
- **Caveat:** QuickSheet is .NET, not Python. But Will boosts *concept* (terminal as first-class UI), not just his ecosystem.

### 5. Charm / Charmbracelet team (@charmcli)
- **Key people:** Christian Rocha (@rechtzeitig, co-founder), Mario Ávila (@muesli, core dev), Ayman Bagabas, Bashbunni.
- **Known for:** Bubble Tea, Lipgloss, Glamour — the Go TUI stack. They retweet novel terminal UX concepts.
- **Platform:** Twitter/X (@charmcli), Discord (charm.sh/chat), YouTube.
- **Strategy:** The "spreadsheet embedded as desktop wallpaper" concept is exactly the kind of creative terminal use they signal-boost. A tweet showing it off with a tag could get amplified.

### 6. Jesse Duffield (@jesseduffield)
- **Known for:** lazygit, lazydocker. Actively signal-boosts other open-source tools.
- **Platform:** Twitter/X.
- **Strategy:** Jesse celebrates creative TUI UX. The "live process output in cells" or "extensions as live data" angle would resonate.

### 7. ThePrimeagen (@ThePrimeagen)
- **Known for:** Vim enthusiast, terminal workflow streamer (YouTube + Twitch). Massive reach (500k+ YouTube subs).
- **Platform:** YouTube, Twitch, Twitter/X.
- **What he covers:** Tools that improve developer workflows in the terminal. He does "underrated TUI tools" videos.
- **Strategy:** Very high reach but hard to target directly. Best path: if QuickSheet hits HN front page, his community often surfaces it. A compelling GIF/demo makes this more likely.
- **Caveat:** He's more likely to feature tools he personally uses (Go/Rust ecosystem bias). .NET is unusual for his audience. But "spreadsheet as wallpaper" is clickbait-level interesting.

---

## Tier 3 — Editorial / Review Sites

### 8. LinuxLinks (Steve Emms)
- **What:** Long-form reviews and "Best of" lists for Linux terminal tools.
- **Editor:** Steve Emms — writes nearly all the TUI/terminal content.
- **How:** No formal submission. Email or be visible enough that he finds it. Articles like "100 Awesome TUI Linux Apps" are regularly updated.
- **Fit:** Good. QuickSheet on Linux with X11 wallpaper mode is exactly his audience.

### 9. LibHunt (libhunt.com/topic/tui)
- **What:** Aggregator that auto-indexes GitHub projects by topic. "Top TUI Projects" page.
- **How:** Mostly automated via GitHub topics. Ensuring QuickSheet has the right topics (`tui`, `terminal`, `spreadsheet`, `cli`) gets it indexed.
- **Fit:** Free, automated. Already covered if topics are set correctly.

### 10. PHP Architect / HowToGeek / itnext.io (roundup writers)
- **What:** Various blogs that write "Best TUI tools" roundups for SEO.
- **How:** No submission path; they find projects via GitHub trending, HN, and awesome-lists.
- **Strategy:** Being on awesome-tuis + TerminalTrove + having a release with binary assets makes discovery by roundup writers more likely.

---

## Implications for QuickSheet

1. **Submit to TerminalTrove immediately** (once a compelling GIF/screenshot exists). This is the single highest-ROI action for ongoing discoverability. Draft already covers submission text; just needs the image asset from the human.

2. **Submit to Console.dev** after TerminalTrove listing. The Console.dev editors check TerminalTrove — being listed there adds social proof. Submission form at console.dev/submit is straightforward.

3. **Ensure GitHub topics are complete.** LibHunt and roundup crawlers use topics. Current topics should include: `tui`, `terminal`, `spreadsheet`, `csv`, `desktop`, `wallpaper`, `dotnet`, `linux`, `windows`, `cli`. Queue a Bucket B action to verify/add missing ones.

4. **Time the HN "Show HN" post to coincide with TerminalTrove + Console.dev features.** If QuickSheet appears on TerminalTrove the same week as a Show HN, the Charm team / Will McGugan / Jesse Duffield are more likely to see it organically. Don't tag them directly (looks spammy) — just be *in the stream* they already monitor.

5. **The GIF/screenshot is the bottleneck.** Every submission channel (TerminalTrove, Console.dev, awesome-tuis, social posts) requires a compelling visual. This is the single blocker the human needs to resolve before any external submission can land.
