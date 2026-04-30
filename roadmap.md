# QuickSheet Commercialization Roadmap

**Date:** 2026-04-29
**Status:** Draft for discussion
**Companion doc:** [report.md](./report.md)

---

## 1. Vision

QuickSheet becomes the **always-visible developer dashboard** — a wallpaper-grade spreadsheet surface that is part scratchpad, part launcher, part live monitor, part AI prompt bar. Single-machine, local-first, keyboard-driven, extensible via the same prefix-cell idiom that already powers `r:` and `i:`.

The bet: hardcore devs already build their own ad-hoc versions of this with bash scripts + Rainmeter + sticky notes + tmux + Raycast. QuickSheet collapses those into one surface that survives reboots and is editable in place.

---

## 2. Target Audience: Hardcore Developers

**Who specifically:**
- Senior ICs, platform/infra engineers, indie hackers
- People who already use tmux, vim, jq, Alfred/Raycast, custom shells
- Comfortable with config-as-text, CLI-first workflows
- Buy paid dev tools without flinching: Sublime Text, Sublime Merge, Charles, Kaleidoscope, Beyond Compare, Soulver, Fork

**Who NOT to target:**
- Mainstream productivity (Notion/Todoist crowd) — wrong UX expectations, want sync
- Mobile users — there is no mobile story
- Teams / enterprise — single-machine by design
- Beginners — keyboard-first, no hand-holding

**Why this audience first:**
- High willingness to pay
- Tolerate rough edges, give precise feedback
- Adopt via Show HN / lobste.rs / r/commandline — channels where QuickSheet's framing already resonates (per `report.md`)
- Word-of-mouth in this segment compounds (one Twitter post from the right person = thousands of installs)

**Positioning line:**
> "Your wallpaper, but it runs commands and remembers things. For developers who hate opening apps."

---

## 3. Monetization

### Recommended model: **Open core + paid Pro + paid extension marketplace**

| Tier | Price | What's in it |
|------|-------|--------------|
| **OSS Core** | Free, MIT | Current feature set: grid, CSV, `r:` / `i:` cells, hyperlinks, autosave, multi-platform desktop embedding |
| **Pro** | $29 one-time, or $39/yr for updates | AI integration (`a:` cell), extension runtime, auto-updater, signed binaries, multi-sheet, formulas (`=`), themes, encrypted cells |
| **Extension Marketplace** | 70/30 split with extension authors | Third-party `we:` weather, `db:` query runners, `gh:` GitHub, `k8s:` cluster, etc. |

**Why this shape:**
- **Open core preserves the supply-chain story** — the trust-building part stays auditable
- **One-time + optional renewal** matches how dev-tool buyers actually behave (Sublime model). Pure subscription will alienate the audience surveyed in `report.md`
- **Marketplace** is the long-term moat: once 50+ extensions exist, switching cost is real
- **Lifetime price anchor** ($29) is below the "no-think" threshold for working devs

### Alternatives considered (and why not)

- **Pure SaaS sub** ($5–10/mo): wrong audience signal, kills the local-first pitch
- **Donation-only**: doesn't fund auto-updates, signing certs, or marketplace infra
- **Enterprise license**: premature; revisit only after individual model proves out
- **Ads in widgets** (Win11 model): would burn the entire trust story documented in `report.md`. Never.

### Revenue sanity check
- 1,000 Pro buyers Y1 @ $29 = $29k — covers Apple/Win signing certs, domain, basic infra, part-time contractor
- 10,000 Pro buyers Y2 @ $29 + $5/yr renewal-attach = $80k+ — sustainable side income
- Target acquisition channels: HN front page (1×/yr realistic), r/commandline + r/Rainmeter (steady drip), word of mouth

---

## 4. AI Integration

### Design principle: **AI is a cell prefix, not a panel**

Same idiom as `r:` / `i:`. New prefix `a:` (or `c:` for Claude). Cell content is the prompt; the cell *below* (or right) receives the response. Multi-cell selection = batch.

```
A1: a: summarize my morning emails
A2: [response streams here, autosaves]

B1: a: what's wrong with this stack trace? {C1::C50}
B2: [response]
C1..C50: <pasted stack trace>
```

### Concrete features

| Feature | Mechanism | Notes |
|---------|-----------|-------|
| `a: <prompt>` cell | Calls Claude API; streams response into adjacent cell | Cell-range refs (`{A1::C10}`) already exist in `CellPrefix.cs` — reuse |
| Selection → "Ask Claude" | Hotkey on multi-select inlines selected cells as context | Pure UI binding |
| Background prompts | Same machinery as `i:` (live subprocess), but the subprocess is the API client | `InlineProcessManager` already handles streaming output |
| BYOK | User pastes Anthropic API key in a Pro settings cell | Stored in OS keyring (DPAPI on Windows, Secret Service on Linux); never in CSV |
| Model picker | Cell-level: `a:claude-opus-4-7: <prompt>` | Trivial parser extension |
| Local model option | Same prefix talks to Ollama / local OpenAI-compatible endpoint | Free Pro feature, removes the "you're sending my notes to Anthropic" objection |

### Why this fits the audience
- No chat window to context-switch into
- AI output is a **first-class cell** — diffable, copyable, export-able to CSV
- Selection-based context = "highlight a stack trace, hit a key, get an explanation pinned to your wallpaper"

### Risks
- **Supply chain**: Claude SDK = NuGet dependency, breaks the zero-deps stance. Mitigation: hand-write the HTTPS POST + SSE parser (Anthropic's API is small enough). Keep it in `Platform/AI/` so the OSS core builds without it.
- **Security**: prompt injection from copied web content lands in cells. Document the risk; don't auto-execute AI-suggested `r:` commands.

---

## 5. Extensibility (`we:` weather, `gh:` GitHub, etc.)

### Design: **Extensions are subprocesses, not plugins**

QuickSheet already has the right mechanism — `r:` runs a command, `i:` streams its output. An extension is a registered prefix that maps to a binary on disk. No DLL loading, no in-process plugin host, no .NET reflection — sidesteps the entire supply-chain risk.

### Mechanism

`%APPDATA%/QuickSheet/extensions/extensions.json`:
```json
{
  "we": { "exec": "qs-weather", "input": "stdin", "output": "stdout", "refresh": 600 },
  "gh": { "exec": "qs-github", "input": "args",  "output": "stdout", "refresh": 60  },
  "k8s":{ "exec": "qs-k8s",     "input": "stdin", "output": "stream", "refresh": 0   }
}
```

Cell `we: 98101` → spawns `qs-weather`, sends `98101` on stdin, displays stdout in cell, refreshes every 600s. Same lifecycle as `i:` cells (uses `InlineProcessManager`).

### Why this is good
- **Authors can write extensions in any language** — Go, Rust, Python, bash. Zero coupling to .NET or C#.
- **Sandboxing is OS-native** — extensions run as the user's process, no new privilege model needed
- **Distribution = a single binary** + a JSON snippet
- **Marketplace just hosts binaries + manifests** — like `npm` for shell tools, except curated and signed
- **Extensions can be tested standalone** — they're just CLI tools

### What we add
- A `qs ext install <name>` CLI command (downloads from marketplace, verifies signature, registers)
- A built-in registry browser (within QuickSheet itself, naturally — a sheet showing installed extensions)
- Code-signed binaries on the marketplace; refuse to install unsigned

### Bootstrap extensions to ship at v1.0
| Prefix | Purpose | Author |
|--------|---------|--------|
| `we:` | Weather (free) | first-party |
| `gh:` | GitHub PR/issue lookup | first-party |
| `cal:` | Calendar today | first-party |
| `clip:` | Clipboard history | first-party |
| `pomo:` | Pomodoro timer | first-party |
| `note:` | Encrypted note (Pro) | first-party |
| `q:` | Run SQL against sqlite | community-friendly example |

---

## 6. Distribution & Auto-Updates

This is the unsexy part that determines whether anyone outside HN ever installs it.

### Required infrastructure

| Item | Cost / yr | Why |
|------|-----------|-----|
| Code-signing cert (Windows EV) | ~$300 | SmartScreen blocks unsigned binaries within hours of any traction |
| Apple Developer account | $99 | Required for macOS notarization (when macOS port lands) |
| Domain + email | ~$30 | quicksheet.dev is available as of this writing |
| GitHub Releases | $0 | Hosts binaries; works as poor-man's CDN at small scale |
| CDN (Cloudflare R2) | ~$0–60 | Switch when GitHub LFS/bandwidth bites |
| Update server | $5/mo VPS | Single Go binary serving a `/latest` JSON manifest is enough |

### Auto-update flow
1. App polls `https://updates.quicksheet.dev/{platform}/{channel}/manifest.json` once per launch + once per day
2. Manifest = `{ version, url, sha256, signature }`
3. App downloads to temp, verifies sha + signature, swaps binary on next launch
4. Channels: `stable`, `beta`, `nightly` — settable in-app
5. Pro users get faster signed updates; OSS users build from source

**Implementation:** small hand-written updater in `Platform/Update/`. No Squirrel.Windows, no AutoUpdater.NET — both are NuGet, both have had supply-chain incidents. ~300 LOC of bespoke code, audit-able.

### Distribution channels (priority order)
1. **GitHub Releases** — primary, day one
2. **Direct download from website** — primary, day one (signed installer)
3. **`winget install quicksheet`** — easy, free, big reach on Windows
4. **`brew install --cask quicksheet`** — when macOS port lands
5. **AUR + flatpak** — Linux power users
6. **Microsoft Store** — last; takes a 12% cut and forces sandboxing that breaks wallpaper embedding
7. **NEVER**: Snap (forced auto-update channel, sandbox issues), Mac App Store (same reasons + sandbox blocks WorkerW-equivalents)

---

## 7. Website (quicksheet.dev)

### Stack: **Static site, hand-written HTML/CSS, hosted on Cloudflare Pages or GitHub Pages**

No Next.js, no Astro, no React. Matches the project's zero-deps ethos. ~5 pages, ~1k LOC of HTML/CSS. Lighthouse 100s.

### Pages
1. **`/`** — hero (animated GIF of typing into wallpaper grid), three use-case strips, install button per OS, "buy Pro" CTA
2. **`/docs`** — single long-page Markdown rendered to static HTML. Searchable client-side.
3. **`/extensions`** — marketplace browser (also static; powered by a JSON manifest in the repo)
4. **`/changelog`** — auto-generated from git tags
5. **`/buy`** — Stripe checkout, license key delivered by email (Stripe → webhook → SES → done; ~50 LOC)
6. **`/blog`** — release announcements, dev diaries (good for HN/SEO)

### Asset list (need to commission or DIY)
- Logo / icon (single SVG, monochrome — terminal-grade aesthetic)
- ~6 short looping screen-capture GIFs (one per use case)
- One hero video (~30s, no narration, just typing)

### Cost: ~$0–20/mo (Cloudflare Pages free tier handles tens of thousands of visits)

---

## Phased Rollout

### Phase 0 — Foundation (1–2 months)
Ship the *credibility surface* before asking for money.
- [ ] Code-signing cert + signed Windows installer
- [ ] Auto-updater (channel: stable only)
- [ ] quicksheet.dev landing page (no purchase yet, just download + GitHub link)
- [ ] `winget` submission
- [ ] First HN post — "Show HN: QuickSheet — vim for spreadsheets, lives on your wallpaper"
- **Goal:** 5,000 GitHub stars, 500 active installs (telemetry: opt-in only)

### Phase 1 — Pro launch (3–4 months in)
- [ ] AI integration (`a:` cell) — biggest single feature
- [ ] Multi-sheet / tabs
- [ ] `=` formulas (basic: arithmetic, ranges, common funcs)
- [ ] Stripe + license keys + Pro gating
- [ ] Encrypted cells
- [ ] macOS port begins (separate track — see Phase 2)
- **Goal:** 1,000 paying Pro users in first 90 days post-launch

### Phase 2 — Extension marketplace (6–9 months in)
- [ ] Extension manifest spec frozen
- [ ] `qs ext install` CLI
- [ ] First-party extensions (`we:`, `gh:`, `cal:`, `clip:`, `pomo:`, `note:`, `q:`)
- [ ] Marketplace site (`/extensions`)
- [ ] Author docs ("write a QuickSheet extension in 10 lines of Python")
- [ ] macOS desktop embedding via NSWindow level hack (native APIs only, no Electron/Tauri)
- **Goal:** 20 community-authored extensions, 3,000 paying users

### Phase 3 — Compounding (year 2)
- [ ] Local model integration (Ollama) for AI-without-cloud
- [ ] Profiles / "loadouts" — one keypress swaps the entire grid (work / personal / oncall)
- [ ] Optional sync via plain `git push` to user's own remote (still local-first)
- [ ] Linux Wayland support (currently X11-only per `CLAUDE.md`)
- [ ] Beta channel + nightly builds
- [ ] First annual user survey

---

## Risks & Tensions

| Risk | Severity | Mitigation |
|------|----------|------------|
| **AI feature breaks "zero NuGet" hard policy** | High | Hand-write Anthropic HTTPS client. Keep AI in `Platform/AI/`, conditionally compiled. OSS users can build without it. |
| **Marketplace supply-chain compromise** | High | Mandatory code signing for marketplace. Quarantine mode for new extensions. Annual key rotation. |
| **Open-core fork undercuts Pro** | Medium | Keep AI / extension runtime / encrypted cells closed. Code-sign Pro binaries. The audience is not the kind that pirates dev tools. |
| **macOS port is huge effort** | Medium | Phase 2; charge a one-time per-OS or bundle. Native-only (no Electron) is non-negotiable. |
| **Auto-update mechanism becomes attack vector** | High | Sign manifests. Pin TLS cert. Audit the updater quarterly. |
| **Audience too narrow to monetize** | Medium | The dev-tool market for $29 one-time products has dozens of profitable examples (Sublime, Beyond Compare, Soulver). Path is well-trodden. |
| **Microsoft / Apple breaks wallpaper embedding** | Low–Medium | Already a risk for Rainmeter / live wallpaper apps and they've survived 15+ years. Maintain both X11 and Wayland paths. |

---

## Open Questions

1. **Pro vs. one-time + paid major versions** (Sublime model)? Lean toward $29 one-time + optional $19 major-version upgrades. Defer subscription discussion.
2. **First-party vs. community extensions for revenue**? First-party only at launch; open community submissions in Phase 2 once moderation tooling exists.
3. **Telemetry**: opt-in or opt-out? Opt-in. Audience hates opt-out; will defect over it.
4. **Company structure**: solo / LLC / open collective? LLC once first dollar lands. Open Collective for community-funded extensions later.
5. **Naming for the Pro tier**: "QuickSheet Pro" vs. "QuickSheet+" vs. just unlocking via license key? Recommend the latter (cleaner messaging).

---

## Immediate Next Actions (this week, if proceeding)

1. Register `quicksheet.dev`
2. File for Windows EV code-signing cert (10-day delivery)
3. Sketch landing page (HTML, no framework) — even before any of the above ships
4. Decide AI integration spec (`a:` vs `c:` prefix; cell layout; key storage)
5. Spec the extension manifest format (write it as a `docs/EXTENSIONS.md` in repo)
6. Open issues in the repo for each Phase 0 item — public roadmap is itself a marketing artifact
