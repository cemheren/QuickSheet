# QuickSheet for dungeon masters

If you're running a TTRPG campaign — D&D, Pathfinder, Mothership, whatever system — **your
desktop wallpaper can be the DM screen.** Initiative order, encounter tables, HP trackers,
dice rolls, and session notes — always visible behind your VTT or PDF, never stealing focus.

> Already using Notion, Obsidian, or a paper DM screen? Think "those, but on the actual
> wallpaper, with live dice and initiative tracking via extensions you can write in 50 lines
> of any language."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/dm-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Your encounter state survives reboots —
the grid is exactly where you left it between sessions.

## What goes on the wallpaper

The starter sheet at `examples/dm-dashboard.csv` gives you a DM screen you can edit in
place:

| Zone              | What it shows                                           | How                                                            |
|-------------------|---------------------------------------------------------|----------------------------------------------------------------|
| **Initiative**    | Sorted combatants, current turn marker, round counter   | [`init`](https://github.com/cemheren/quicksheet-init-ext)      |
| **Dice**         | Roll results — `2d6+3`, `1d20`, `4d8` etc.             | [`roll`](https://github.com/cemheren/quicksheet-roll-ext)      |
| **Party HP**      | Character names, current/max HP, status conditions      | Plain cells — edit during combat                               |
| **Encounter**     | Monster stat blocks, AC, HP, attacks                    | Plain cells — paste from your prep                             |
| **Session notes** | What happened, NPC names, plot hooks                    | Plain cells — type as you go                                   |

Both extensions install with a single `ext: github:cemheren/quicksheet-<name>` cell.
QuickSheet clones the repo, starts the subprocess, and the prefix is live.

## Roll dice without leaving the grid

The `roll:` extension understands standard dice notation:

```
roll: 1d20+5       → "18"
roll: 2d6+3        → "11"
roll: 4d6kh3       → "14"  (keep highest 3 — ability score gen)
roll: 1d100        → "73"  (percentile)
```

Put a `roll:` cell next to each combatant for attack rolls, or keep a scratch column
for quick checks.

## Track initiative in real time

The `init:` extension manages turn order:

```
init: add Wizard 18, Fighter 22, Goblin 14, Goblin 9
init: next              → advances turn marker
init: sort              → re-sort by initiative value
init: remove Goblin 14  → drops a dead combatant
```

One cell drives the tracker. The output shows the current combatant highlighted and
the round counter.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Nord, Dracula, Synthwave, Gruvbox, Matrix. Dracula
and Synthwave fit the "gaming night" vibe; Matrix for cyberpunk campaigns.

## Why wallpaper > a browser tab

- **Always visible** — no Alt+Tab to "check the initiative." It's right there when you
  glance past your VTT window.
- **Survives reboots** — autosaved CSV opens to the same encounter state next session.
- **No SaaS** — your encounter data is a plain CSV file. No accounts, no subscriptions.
- **Click-through** — the grid sits behind all windows. Your VTT, PDFs, and Discord stay
  on top. You see the DM screen in the gaps.
- **Instant edits** — click a cell, type the new HP. Done. No forms, no modals.
- **Extensions are scripts** — write a 50-line extension for your homebrew system.

## Recipe: combat encounter dashboard

1. **Row 1:** Initiative cell (`init: add ...`) + round counter.
2. **Rows 2–8:** One row per combatant — Name, AC, HP current, HP max, Status.
3. **Column to the right:** `roll:` cells for attack rolls and damage.
4. **Bottom strip:** Session notes, loot, XP tally (plain cells, Σ for column sums).

When combat starts, type initiative values into the `init:` cell. Advance turns with
`init: next`. Cross off HP as damage lands. The whole encounter fits on the wallpaper.

## Build your own for homebrew systems

The extension protocol is two message types
([extension-protocol.md](extension-protocol.md)). A 50-line Python script can:
- Roll with custom dice (Fate dice, exploding dice, dice pools)
- Look up monsters from a local JSON bestiary
- Track spell slots or class resources

Zero dependencies required. `ext: github:yourname/quicksheet-myext` and it's live.

## Share your setup

Post your DM screen wallpaper on r/DnD, r/dndnext, r/rpg, r/unixporn (the rice angle
is real for TTRPG nerds who also rice). Open an issue if you want a new prefix
officially listed.
