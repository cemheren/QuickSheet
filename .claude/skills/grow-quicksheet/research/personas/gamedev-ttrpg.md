# Persona 5 — Indie game devs + TTRPG game masters

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- Two adjacent subsegments, **same wallpaper-mode-loving culture**: indie game devs (Godot/Unity solo + small-team) and tabletop RPG GMs (D&D 5e/Pathfinder/PbtA/custom). US/EN active: ~300k+ indie devs, ~10M people who have GM'd at least once.
- Already on r/unixporn-adjacent boards. They make custom dev tools and GM screens *because they enjoy it.* Highest evangelism rate of any persona — they post their setups voluntarily.
- Highest-leverage wallpaper cells: per-system status (build server / playtest queue / itch.io revenue / wishlist count) for devs; campaign tracker (initiative, HP, NPC pile, session notes, loot table) for GMs.
- Where to seed: r/gamedev, r/IndieDev, r/godot, r/Unity3D, r/DMAcademy, r/rpg, r/osr, r/RPGtools, r/unixporn (rice angle), itch.io devlog network, Roll20 forums.
- Direct competition: dev side — none (Steamworks is browser). GM side — Roll20 (cloud, $$), Foundry VTT ($, electron), printed sheets. **Nobody offers an offline desktop-layer.**

## 1. Profile

### 1a. Indie game devs

- Solo + 2–5-person studios; mix of Godot 4 (rising fast), Unity (still dominant), Unreal, custom engines. Plus narrative-IF/Twine writers and pixel-art game devs.
- Hardware: gaming desktop or beefy laptop. Often dual-monitor (engine + reference/Trello).
- Lifecycle: long pre-launch period, brief Steam/itch launch window, perpetual patch-and-engage cycle. Wishlist count is *the* number.
- Buying brain: love free + open-source on principle; will install something because the README screenshot is cool.

### 1b. TTRPG GMs

- Range from D&D 5e (mass market, ~50M ever-played, ~3M active GMs) to OSR/indie systems (~50k-100k zealous, very online crowd).
- Hardware: laptop at the table or webcam on remote (Discord + Roll20/Foundry). **One screen they want hidden from players.**
- Buying brain: tactile, customisation-first, anti-SaaS for player data, will literally hand-write tools for one campaign.
- Cultural moment: post-OGL-fiasco (Jan 2023) backlash against WotC + Roll20 dependencies. Open-source/local-first tooling has unusually high goodwill.

## 2. Why their desktop is wasted

### Devs

- Steamworks/itch dashboards behind a browser tab they have to refocus to.
- Trello/Notion/Jira behind tabs.
- Discord pings.

What QuickSheet replaces: the *intent* of opening 6 browser tabs to "check on the launch." Wallpaper shows wishlist, latest Steam review, build status, today's revenue — *no clicks*.

### GMs

- Session prep notes in 14 Google Docs tabs.
- A PDF reader showing the rulebook.
- A spreadsheet of NPCs they keep forgetting to alt-tab to.

What QuickSheet replaces: the GM screen behind the laptop lid. With the wallpaper, every NPC HP / initiative slot / loot pool / session-XP cell is *always visible* on the GM-side monitor while players see only the shared map. **The GM laptop's screen is the perfect QuickSheet target: never shown to players, always glanceable to the GM.**

## 3. Glanceable data they actually want behind windows

### Devs

1. **Wishlist count** + 7-day delta.
2. **Today's revenue** (Steam + itch + Patreon combined).
3. **Latest review** (most recent Steam review snippet + score).
4. **Build status** — last CI run for the engine project.
5. **Playtester queue** — names + last-played date.
6. **Bug tracker count** (or just labeled-issue count from GitHub).
7. **Devlog streak** — days since last itch devlog post.

### GMs

1. **Initiative tracker** — current turn highlighted, HP next to each name.
2. **Active NPC pile** — name, role, motivation, secret, current HP.
3. **Quest log** — per-quest status (open / in-progress / done).
4. **Loot table** — items the party has, by character.
5. **Session timer + break countdown** — "next break in 23 min."
6. **Inspiration / luck dice pool** — single cell, count.
7. **Random-encounter trigger** — `i: roll d100 | match table` cell.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                    | Type | Cost | Hit prob | Why                                                                             |
|------|-----------------------------------------|------|------|----------|----------------------------------------------------------------------------------|
| 1    | `steam:` wishlist + revenue lookup      | ext  | med  | high     | Steamworks Partner API; auth via token. Hits every Steam dev.                   |
| 2    | `itch:` itch.io devlog/revenue           | ext  | low  | high     | Public stats per game; no auth for public views. Indie+jam crowd lives there.   |
| 3    | `roll:` dice roller + table lookup       | ext  | low  | high     | `roll: 4d6 drop lowest` etc. + CSV table lookups for encounter tables.          |
| 4    | `init:` initiative tracker               | ext  | low  | high     | Active-turn highlight, HP edits inline. GM-side wallpaper headline.             |
| 5    | **Value-driven cell colour**             | feat | low  | high     | HP at 0 = red, low = yellow. Same feature, more personas hit.                   |
| 6    | **Per-cell ticking timer**               | feat | low  | high     | Session timer + break countdown — recurring across personas.                    |
| 7    | `npc:` NPC sheet template                | ext  | low  | med-high | Pre-built CSV with name/race/motivation/secret columns + load command.           |
| 8    | `5esrd:` D&D 5e SRD spell/monster lookup | ext  | med  | high     | Open5e API; perfect cell-by-cell lookup for spells/monsters during session.     |
| 9    | `patreon:` patron count                  | ext  | med  | med      | OAuth-heavy; defer.                                                              |
| 10   | `godot:` build/test status from .godot   | ext  | high | med      | Engine-specific; small but engaged.                                              |

Features 5 and 6 are the recurring "trinity" needed by SREs/lawyers/accountants/artists *and* this persona. Highest cross-persona leverage in the whole research phase.

## 5. Where they hang out (desktop-customization first)

- **r/unixporn** — both subgroups overlap with rice culture. GM-screen wallpaper + dev-stat wallpaper are *both* viable rice posts.
- **r/IndieDev (250k), r/gamedev (1.5M), r/godot (200k), r/Unity3D (340k), r/UnrealEngine, r/pcgaming dev-thread**.
- **r/DMAcademy (550k), r/rpg (1.6M), r/dndnext, r/osr, r/RPGtools, r/dungeonmasters**.
- **itch.io devlog network** — devs read each other's devlogs religiously; a "I built this for my own dev process" devlog ranks well.
- **Discords:** Godot, Unity, Roll20, Foundry VTT communities, individual TTRPG creator servers (MCDM, Mörk Borg, Sly Flourish).
- **Podcasts/blogs:** Sly Flourish (Mike Shea), The Angry GM, Justin Alexander's The Alexandrian, Chris Zukowski's *How to Market a Game*.
- **People to be visible to:** Mike Shea (Sly Flourish, lazy DM advocate, would love a low-prep tool), Chris Zukowski (gamedev marketing analyst), itch.io's @leafo. Plus Godot evangelists like Juan Linietsky.

## 6. Discoverability hooks

Hero images — **two different ones**, because this is two personas under one paper:

- **Dev hero**: Godot or Unity open with a level scene; QuickSheet wallpaper visible to the right showing wishlist count, today's revenue, latest review snippet, build status. Title: "I track my Steam wishlist and itch revenue on my wallpaper now."
- **GM hero**: Discord call window open with players; on the GM's screen behind it, QuickSheet wallpaper with initiative order, NPC HP grid, current quest, loot table. Title: "My DM screen is a wallpaper now (with dice roller, no Roll20)."

Headlines that land:

- "I built a free, local-first DM screen for D&D — your wallpaper as your GM tools"
- "Steam wishlist + itch revenue on my desktop wallpaper (zero deps, MIT)"
- "Roll20 was eating my soul. Now my GM tools are 200 lines on my wallpaper."

The OGL backlash + Roll20-fatigue is a real conversation hook for the GM half. **Frame as "your campaign data stays on your laptop."**

Avoid: TUI screenshots, terminal demos. The TTRPG audience will tune out instantly.

## 7. Implications (queue these)

1. **Build `roll:` extension first** (Bucket F, post-research). Lowest cost, broadest GM appeal, screenshot-friendly (`roll: 4d6 drop lowest = [5,4,3,2] = 12`). Bonus: also useful for non-gamers as a quick die-roller.
2. **Build `init:` initiative tracker extension** second. The single most-requested DM-tool feature on r/DMAcademy/r/RPGtools survey threads. Pairs with `roll:`.
3. **Build `itch:` extension** for the dev half. Low auth burden, hits the audience that *will* devlog about discovering QuickSheet.
4. **Bucket A: `examples/dm-screen.csv` + `examples/gamedev-launch.csv`** — pre-built starter sheets. Zero-LOC onboarding, maximum showcase. Pair with `docs/for-dms.md` and `docs/for-indiedevs.md`.
5. **Bucket C: r/unixporn rice post** with GM-screen screenshot. *Most-likely-to-go-viral* persona — single screenshot, novel use, strong "I'd never thought of that" reaction.

Cross-link:
- [developers-sre.md](developers-sre.md) — same colour/timer features.
- [artists-visual.md](artists-visual.md) — same hex-color cell ask (GMs colour-code factions).
