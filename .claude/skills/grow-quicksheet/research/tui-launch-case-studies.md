# TUI Launch Case Studies: Zero to Thousands of Stars

## TL;DR

- Average "Show HN" launch yields **121 stars/24h, 289/week** (arxiv:2511.04453, n=138).
- **Timing matters more than the "Show HN" tag**: 12:00–17:00 UTC is the golden window; **Sunday 11–16 UTC** has the highest breakout rate (~15.7%).
- Every breakout TUI had **one GIF or screenshot that sold the tool in 3 seconds**.
- The creator's presence in the launch thread is a multiplier — lazygit, harlequin, superfile all had founders replying within minutes.
- Post-launch: sustained growth comes from **weekly release cadence**, not one-shot virality.

---

## Case Studies

### 1. lazygit — 79k★ (Go, 2018)

- **Pain point**: Git CLI intimidation. One-screen visual for staging, branching, rebasing.
- **Launch**: "Show HN: lazygit – simple terminal UI for git commands" (Oct 2018, HN item 18354193).
- **First week**: ~1,000+ stars. Front-paged HN within hours.
- **What worked**:
  - Demo GIF at top of README showing a real staging → commit → push flow.
  - Cross-platform Go binary — `brew install lazygit` one-liner.
  - Jesse Duffield answered every HN comment personally.
  - Title was specific and curiosity-provoking (not "A new tool" — literally "simple terminal UI for git commands").
- **Sustained growth**: Weekend side-project cadence, frequent releases, eventually 79k stars and handed off maintenance. Active contributor community formed organically.
- **Lesson for QuickSheet**: The title "simple terminal UI for git" told you what it does and why in 7 words. QuickSheet's equivalent: "spreadsheet that lives on your desktop wallpaper."

### 2. btop — 24k★ (C++, 2021)

- **Pain point**: htop is ugly and limited. btop is a resource monitor that looks like a dashboard.
- **Launch**: "Show HN: btop – Resource monitor" (HN item 28815921).
- **What worked**:
  - **Visuals sold it.** The README had 3 themed screenshots (Dracula, Nord, default) showing CPU/memory/disk/network in one glance. People starred it before even installing.
  - Evolution story: bashtop (Python) → bpytop (Python, faster) → btop++ (C++, fastest). Each rewrite was a fresh launch opportunity.
  - Multi-platform from day one.
- **Lesson for QuickSheet**: Theme screenshots are launch ammunition. Each theme preset is a separate screenshot opportunity that appeals to r/unixporn and ricing communities.

### 3. Harlequin — 6k★ (Python, 2023)

- **Pain point**: No good terminal SQL IDE. DBeaver is heavy; sqlite3 CLI is primitive.
- **Launch**: Show HN, ~2023. "The SQL IDE for your terminal."
- **What worked**:
  - One-liner pitch: "DBeaver for the terminal, but fast."
  - Extensible adapter system (DuckDB, PostgreSQL, MySQL, BigQuery) — each adapter was a discovery surface.
  - Creator Ted Conbeer was active on Python podcasts post-launch.
  - Frequent minor releases kept it appearing in GitHub Trending.
- **Star growth**: Steady climb to 6k, not a single spike. Python Weekly, podcast coverage, and adapter PRs from community.
- **Lesson for QuickSheet**: Extension ecosystem = discovery surface. Each `quicksheet-*-ext` repo is a potential inbound link. Harlequin proved that adapters/plugins multiply visibility.

### 4. superfile — 17k★ (Go, 2024)

- **Pain point**: Terminal file managers (ranger, nnn, lf) are functional but ugly.
- **Launch**: Show HN + Reddit (r/commandline, r/linux, r/programming) simultaneously.
- **What worked**:
  - **Visual shock factor**: A terminal file manager that looks like a GUI app. Multi-panel, icons, Dracula/Nord themes, smooth animations. Screenshots were inherently shareable.
  - **Origin story**: Created by a high schooler in Taiwan. Human interest angle amplified sharing.
  - Cross-posted to multiple subreddits; each post hit front page independently.
  - One-liner install: `bash -c "$(curl -sLo- https://superfile.netlify.app/install.sh)"`.
- **17k stars in <1 year** — fastest-growing file manager TUI ever.
- **Lesson for QuickSheet**: The "wallpaper spreadsheet" angle is QuickSheet's visual shock factor. A single screenshot of a beautiful wallpaper-embedded grid with live data would be the superfile-equivalent moment.

### 5. yazi — 33k★ (Rust, 2023–2024)

- **Pain point**: ranger/nnn are slow on large directories. No image previews.
- **Launch**: Show HN, multiple front-page appearances.
- **What worked**:
  - **Technical differentiation**: True async I/O (no UI freeze), image preview in terminal (Kitty/Sixel).
  - Lua-based plugin system + package manager (`ya pkg`) — community could extend without forking.
  - Vim-style keybindings attracted the power-user audience.
  - Repeated HN front-page appearances on major releases (not just launch day).
- **33k stars** by mid-2024.
- **Lesson for QuickSheet**: Re-launches work. Each major feature (extensions, sparklines, inline processes) is a valid "Show HN v2" moment. Plan for multiple launch windows, not one shot.

### 6. atuin — 30k★ (Rust, 2023)

- **Pain point**: Shell history is ephemeral, unsearchable, lost across machines.
- **Launch**: Show HN + word-of-mouth.
- **What worked**:
  - **One-liner value prop**: "Magical shell history with sync, search, and stats."
  - E2E encrypted sync was a trust differentiator.
  - Ellie Huxtable (creator) was active in HN comments and on Twitter/fediverse.
  - Frictionless install (curl | bash one-liner).
- **Lesson for QuickSheet**: "Sync" and "encrypted" are power words in launch titles. QuickSheet's analogue: "live desktop data" or "always-visible spreadsheet."

---

## Quantitative Patterns (from arxiv:2511.04453)

| Metric | Average | Top-quartile |
|--------|---------|-------------|
| Stars in 24h | 121 | 300+ |
| Stars in 48h | 189 | 500+ |
| Stars in 7d | 289 | 1000+ |
| Optimal post window | 12:00–17:00 UTC | Sunday 11–16 UTC (15.7% breakout) |

- "Show HN" tag provides **no statistical edge** over regular HN posts when controlling for timing and engagement.
- The **first 30 minutes** are critical: 8–10 upvotes + 2–3 genuine comments in the first half-hour predict front-page.
- **Show HN is getting crowded**: 30 → 65 weekly launches (2023→2025). Differentiation matters more than timing.

## Cross-Cutting Success Factors

| Factor | lazygit | btop | harlequin | superfile | yazi | atuin |
|--------|---------|------|-----------|-----------|------|-------|
| Demo GIF in README | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| One-liner install | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Creator in HN thread | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Multi-platform | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Plugin/extension system | ❌ | ❌ | ✅ | ✅ | ✅ | ❌ |
| Repeated launches | ✅ | ✅ | ❌ | ❌ | ✅ | ❌ |
| Human interest angle | ❌ | ❌ | ❌ | ✅ | ❌ | ❌ |

**6/6 had**: demo GIF, one-liner install, founder in comments, multi-platform.

---

## Implications for QuickSheet

### 1. The single most important pre-launch task is a demo GIF
Every breakout TUI had one. QuickSheet's README currently has static screenshots. A 15-second GIF showing: open wallpaper → type in a cell → see live extension output → sparkline updates — would be the single highest-ROI asset. **This blocks on human capture** but should be priority #1.

### 2. Plan for multiple launch windows, not one shot
yazi and lazygit both re-launched on HN for major version bumps. QuickSheet should plan:
- **Launch 1**: "Show HN: A spreadsheet that lives on your desktop wallpaper" (core product)
- **Launch 2**: "Show HN: 30 extensions turn your desktop wallpaper into a dashboard" (extension ecosystem)
- **Launch 3**: Niche-targeted posts (r/selfhosted for homelab, r/unixporn for ricing, r/DMAcademy for TTRPG)

### 3. Optimal posting: **Sunday 12:00 UTC**
Data-backed best time. Prepare the post Saturday night, publish Sunday morning UTC.

### 4. Extension repos are Harlequin-style discovery surfaces
Each `quicksheet-*-ext` repo should have its own polished README with install one-liner and screenshot. When someone discovers `quicksheet-stock-ext` via GitHub search for "stock price terminal," they find QuickSheet.

### 5. One-liner install is table stakes
QuickSheet currently requires `dotnet build` from source. A `curl | bash` installer or a single-binary GitHub Release download would dramatically lower the barrier. The existing v0.36.0 release helps, but the README install section should lead with the easiest path.

---

## Sources

- [HN: lazygit launch (item 18354193)](https://news.ycombinator.com/item?id=18354193)
- [HN: btop launch (item 28815921)](https://news.ycombinator.com/item?id=28815921)
- [arxiv:2511.04453 — Launch-Day Diffusion](https://arxiv.org/abs/2511.04453)
- [Myriade: Best Time to Post Show HN](https://www.myriade.ai/blogs/when-is-it-the-best-time-to-post-on-show-hn)
- [Flowjam: Front Page of HN in 2025](https://www.flowjam.com/blog/how-to-get-on-the-front-page-of-hacker-news-in-2025-the-complete-up-to-date-playbook)
- [lazygit blog](https://jesseduffield.com/)
- [superfile GitHub](https://github.com/yorukot/superfile)
- [yazi GitHub](https://github.com/sextants/yazi)
- [harlequin GitHub](https://github.com/tconbeer/harlequin)
- [atuin GitHub](https://github.com/atuinsh/atuin)
